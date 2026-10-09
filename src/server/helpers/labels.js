const labelUtils = {
  /**
   * Whether a cell is the ID label for a sheet type.
   * @param {*} cell
   * @param {string} sheetType
   * @returns {boolean}
   */
  isSheetTypeCell: function (cell, sheetType) {
    if (typeof cell !== "string" || !sheetType) {
      return false;
    }
    return (
      new RegExp(sheetType, "i").test(cell) &&
      /\bID\b/i.test(cell) &&
      cell.indexOf("script") === -1 &&
      cell.indexOf("More IDs are available") === -1
    );
  },

  /**
   * Finds a sheet type's ID in an IDS tab.
   * @param {string} spreadsheetId
   * @param {string} sheetName
   * @param {string} sheetType
   * @param {Array<Array<*>>} [values]
   * @returns {string} "" when not found.
   */
  findSheetTypeID: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues,
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`,
        );
        return null;
      }
      values = batchResult[0].values;
    }

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (labelUtils.isSheetTypeCell(values[i][j], sheetType)) {
          var cellA1 = sheetRefs.columnToLetter(j + 2) + (i + 1);
          var accessA1 = sheetRefs.columnToLetter(j + 4) + (i + 1);
          var importedA1 = sheetRefs.columnToLetter(j + 4) + (i + 2);

          var accessValue = "";
          var importValue = "";

          if (values[i] && values[i][j + 3]) {
            accessValue = values[i][j + 3];
          }
          if (values[i + 1] && values[i + 1][j + 3]) {
            importValue = values[i + 1][j + 3];
          }

          return {
            id: values[i][j + 2],
            cell: {
              row: i + 1,
              col: j + 2,
              range: sheetName + "!" + cellA1,
            },
            accessStatus: {
              row: i + 1,
              col: j + 4,
              range: sheetName + "!" + accessA1,
              value: accessValue,
            },
          };
        }
      }
    }
    return null;
  },

  /**
   * Finds a sheet type's row in an IDS tab: id, template and version.
   * @param {string} spreadsheetId
   * @param {string} sheetName
   * @param {string} sheetType
   * @param {Array<Array<*>>} [values]
   * @returns {Object|null}
   */
  findSheetTypeURL: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues,
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`,
        );
        return null;
      }
      values = batchResult[0].values;
    }

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (labelUtils.isSheetTypeCell(values[i][j], sheetType)) {
          var versionA1 = sheetRefs.columnToLetter(j + 6) + (i + 1);
          var templateA1 = sheetRefs.columnToLetter(j + 1) + (i + 2);
          var oldVersionA1 = sheetRefs.columnToLetter(j + 7) + (i + 1);

          var versionValue = "";
          var oldVersionValue = "";
          if (values[i] && values[i][j + 5]) {
            versionValue = values[i][j + 5];
          }
          if (values[i] && values[i][j + 6]) {
            oldVersionValue = values[i][j + 6];
          }

          return {
            id: values[i][j + 2],
            template: {
              row: i + 2,
              col: j + 1,
              range: sheetName + "!" + templateA1,
            },
            version: {
              row: i + 1,
              col: j + 6,
              range: sheetName + "!" + versionA1,
              value: versionValue,
            },
            oldVersion: {
              row: i + 1,
              col: j + 7,
              range: sheetName + "!" + oldVersionA1,
              value: oldVersionValue,
            },
          };
        }
      }
    }
    return null;
  },

  /**
   * Finds a sheet type's template ID in an IDS tab.
   * @param {string} sheetID
   * @param {string} sheetName
   * @param {string} sheetType
   * @returns {string|null}
   */
  findSheetTemplateID: function (sheetID, sheetName, sheetType) {
    try {
      console.log(
        `Finding template ID for sheet: ${sheetID}, sheet name: ${sheetName}, type: ${sheetType}`,
      );

      var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, sheetID);
      if (!spreadsheet) {
        console.log(`Could not access spreadsheet with ID: ${sheetID}`);
        return null;
      }

      var sheet = SheetsAPI.getSheetByName(spreadsheet, sheetName);
      if (!sheet) {
        console.log(`Could not find sheet: ${sheetName}`);
        return null;
      }

      var formulas = SheetsAPI.batchGetFormulas(sheetID, [sheetName]);
      var values = SheetsAPI.batchGetValues(sheetID, [sheetName]);

      if (!formulas || !formulas[0] || !formulas[0].values) {
        console.log(`Could not fetch formulas from sheet: ${sheetName}`);
        return null;
      }

      if (!values || !values[0] || !values[0].values) {
        console.log(`Could not fetch values from sheet: ${sheetName}`);
        return null;
      }

      var formulaData = formulas[0].values;
      var valueData = values[0].values;

      console.log(
        `Searching ${formulaData.length} rows for template HYPERLINK formulas`,
      );

      var templateID = null;
      var version = null;

      for (var i = 0; i < formulaData.length; i++) {
        for (var j = 0; j < formulaData[i].length; j++) {
          var formula = formulaData[i][j];

          if (
            formula &&
            typeof formula === "string" &&
            formula.toUpperCase().includes("HYPERLINK") &&
            formula.toLowerCase().includes("copy")
          ) {
            console.log(
              `Found potential template link in row ${i + 1}, col ${
                j + 1
              }: ${formula}`,
            );

            var templateUrl = sheetRefs.extractUrlFromHyperlink(formula);
            if (templateUrl) {
              templateID = sheetRefs.extractSheetId(templateUrl);
              if (templateID) {
                console.log(`Found template ID: ${templateID}`);
              }
            }
          }
        }
        if (templateID) {
          break;
        }
      }

      if (templateID) {
        var currentSheetVersionInfo = versionUtils.findSheetVersion(
          sheetID,
          sheetName,
          sheetType,
          valueData,
        );
        if (currentSheetVersionInfo && currentSheetVersionInfo.latestVersion) {
          console.log(
            `Template version (from latest): ${currentSheetVersionInfo.latestVersion}`,
          );
          return {
            templateID: templateID,
            templateVersion: currentSheetVersionInfo.latestVersion,
          };
        }
      }

      console.log(
        `No template HYPERLINK with "copy" found in sheet: ${sheetName}`,
      );
      return null;
    } catch (error) {
      errors.report("labelUtils.findSheetTemplateID", error, {
        note: `Error finding template ID`,
        sheetID: sheetID,
        sheetName: sheetName,
        sheetType: sheetType,
      });
      return null;
    }
  },

  /**
   * Appends the IDS Master ID writes to a batch update.
   * @param {Array<Object>} batchUpdate Mutated in place.
   * @param {string} sheetType
   * @param {string} newSheetID
   * @param {Array<Array<*>>} idsData
   * @param {string} idMasterID
   * @returns {Array<Object>} The same batch.
   */
  addIDUpdatesToBatch: function (
    batchUpdate,
    sheetType,
    newSheetID,
    idsData,
    idMasterID,
  ) {
    try {
      if (newSheetID && idMasterID) {
        var thisSheetInfo = labelUtils.findSheetTypeID(
          newSheetID,
          "IDS",
          "This Sheet ID",
          idsData,
        );
        var idMasterInfo = labelUtils.findSheetTypeID(
          newSheetID,
          "IDS",
          "IDS Master's",
          idsData,
        );

        if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
          batchUpdate.push({
            range: thisSheetInfo.cell.range,
            values: [[newSheetID]],
          });
        }
        if (idMasterInfo && idMasterInfo.cell && idMasterInfo.cell.range) {
          batchUpdate.push({
            range: idMasterInfo.cell.range,
            values: [[idMasterID]],
          });
        }
      }
      return batchUpdate;
    } catch (error) {
      errors.report("labelUtils.addIDUpdatesToBatch", error, {
        note: `Error adding ID updates to batch`,
        batchUpdate: batchUpdate,
        sheetType: sheetType,
        newSheetID: newSheetID,
        idsData: idsData,
        idMasterID: idMasterID,
      }, errors.CODES.RECOVERED);
      return batchUpdate;
    }
  },
};
