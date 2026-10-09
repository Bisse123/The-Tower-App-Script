const workshopWriter = {

  /**
   * Builds the batch update that writes WorkshopLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldWorkshopLevels
   * @param {Object} oldWorkshopPlusLevels
   * @param {*} hasPresets
   * @param {Object} masterSheetData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateWorkshopLevels: function (
    sheetName,
    oldWorkshopLevels,
    oldWorkshopPlusLevels,
    hasPresets,
    masterSheetData,
  ) {
    try {
      console.log("Called: workshopWriter.updateWorkshopLevels");
      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];

      var upgradeCol = 0;
      var enhancementCol = 0;

      for (var i = 0; i < headerRow.length; i++) {
        var cellValue = String(headerRow[i] || "").toLowerCase();
        if (cellValue.includes("workshop upgrade")) {
          upgradeCol = i + 1;
        } else if (cellValue.includes("workshop enhancement")) {
          enhancementCol = i + 1;
        }
        if (upgradeCol > 0 && enhancementCol > 0) {
          break;
        }
      }

      if (upgradeCol === 0) {
        console.log(`Workshop Upgrade column not found`);
        return {
          success: false,
          message: `Workshop Upgrade column not found`,
        };
      }
      if (enhancementCol === 0) {
        console.log(`Workshop Enhancement column not found`);
        return {
          success: false,
          message: `Workshop Enhancement column not found`,
        };
      }

      var workshopUnlocked = [];
      var workshopLevels = [];
      var workshopPlusLevels = [];
      var workshopLevelsStartCol = upgradeCol + 1;
      var workshopLevelsEndCol =
        workshopLevelsStartCol + oldWorkshopLevels.presetNames.length * 2 - 1;
      var workshopPlusLevelsStartCol = enhancementCol + 2;
      var workshopPlusLevelsEndCol =
        workshopPlusLevelsStartCol +
        oldWorkshopPlusLevels.presetNames.length -
        1;
      for (var i = 2; i < masterSheetData.length; i++) {
        var row = masterSheetData[i];

        var workshopName = row[upgradeCol - 1];
        if (workshopName && workshopName.startsWith("=")) {
          workshopName = workshopName
            .split(",")[1]
            .trim()
            .replace(/['"()]/g, "");
        }
        if (workshopName && oldWorkshopLevels.data.hasOwnProperty(workshopName)) {
          var oldWorkshopRowData = oldWorkshopLevels.data[workshopName];
          workshopUnlocked.push([oldWorkshopRowData.unlocked || ""]);
          workshopLevels.push(oldWorkshopRowData.levels || []);
        } else {
          workshopUnlocked.push([""]);
          workshopLevels.push([]);
        }
        var enhancementName = row[enhancementCol - 1];
        if (enhancementName && enhancementName.startsWith("=")) {
          var parts = enhancementName.split(",");
          if (parts.length === 4) {
            enhancementName = parts[2].trim();
          } else {
            enhancementName = parts[parts.length - 1].trim();
          }
          enhancementName = enhancementName.replace(/['"()]/g, "");
        }
        if (enhancementName && oldWorkshopPlusLevels.data.hasOwnProperty(enhancementName)) {
          var enhancementData = oldWorkshopPlusLevels.data[enhancementName];
          workshopPlusLevels.push(enhancementData);
        } else {
          workshopPlusLevels.push([]);
        }
      }

      var batchUpdate = [];
      var startRow = 3;
      if (upgradeCol > 1 && workshopUnlocked.length > 0) {
        var unlockedCol = sheetRefs.columnToLetter(upgradeCol - 1);
        var unlockedRange = `${sheetName}!${unlockedCol}${startRow}:${unlockedCol}${
          workshopUnlocked.length + startRow - 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: workshopUnlocked,
        });
      }

      if (upgradeCol > 0 && workshopLevels.length > 0) {
        var upgradeHeaders = [];
        oldWorkshopLevels.presetNames.forEach(function (presetName) {
          upgradeHeaders.push(presetName, "");
        });
        var levelsStartCol = sheetRefs.columnToLetter(workshopLevelsStartCol);
        var levelsEndCol = sheetRefs.columnToLetter(workshopLevelsEndCol);
        var levelsRange = `${sheetName}!${levelsStartCol}${startRow}:${levelsEndCol}${
          workshopLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: levelsRange,
          values: workshopLevels,
        });
        if (hasPresets) {
          var levelsHeaderRange = `${sheetName}!${levelsStartCol}1:${levelsEndCol}1`;
          batchUpdate.push({
            range: levelsHeaderRange,
            values: [upgradeHeaders],
          });
        }
      }

      if (enhancementCol > 0 && workshopPlusLevels.length > 0) {
        var plusHeaders = [];
        oldWorkshopPlusLevels.presetNames.forEach(function (presetName) {
          plusHeaders.push(presetName);
        });
        var plusStartCol = sheetRefs.columnToLetter(workshopPlusLevelsStartCol);
        var plusEndCol = sheetRefs.columnToLetter(workshopPlusLevelsEndCol);
        var plusRange = `${sheetName}!${plusStartCol}${startRow}:${plusEndCol}${
          workshopPlusLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: plusRange,
          values: workshopPlusLevels,
        });
        if (hasPresets) {
          var plusHeaderRange = `${sheetName}!${plusStartCol}1:${plusEndCol}1`;
          batchUpdate.push({
            range: plusHeaderRange,
            values: [plusHeaders],
          });
        }
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Workshop levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for workshop levels`,
      };
    } catch (error) {
      var errorReport = errors.report("workshopWriter.updateWorkshopLevels", error, {
        sheetName: sheetName,
        oldWorkshopLevels: oldWorkshopLevels,
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
        hasPresets: hasPresets,
        masterSheetData: masterSheetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes WorkshopPlusRatios into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldWorkshopPlusRatios
   * @param {Object} desiredRatiosData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateWorkshopPlusRatios: function (
    sheetName,
    oldWorkshopPlusRatios,
    desiredRatiosData,
  ) {
    try {
      console.log("Called: workshopWriter.updateWorkshopPlusRatios");
      if (!desiredRatiosData) {
        console.log(`Not enough data in Desired Ratios sheet`);
        return {
          success: false,
          message: `Not enough data in Desired Ratios sheet`,
        };
      }

      var headerRow = desiredRatiosData[0];
      var workshopEnhancementNameCol = headerRow.indexOf(
        "Workshop Enhancement",
      );

      if (workshopEnhancementNameCol === -1) {
        console.log(`Workshop Enhancement column not found`);
        return {
          success: false,
          message: `Workshop Enhancement column not found`,
        };
      }

      var ratiosToUpdate = [];
      var batchUpdate = [];
      for (var i = 1; i < desiredRatiosData.length; i++) {
        var row = desiredRatiosData[i];
        var workshopPresetIndex = row.indexOf("Workshop preset");
        if (workshopPresetIndex !== -1) {
          var presetValue = oldWorkshopPlusRatios["Workshop preset"] || null;
          var presetCol = sheetRefs.columnToLetter(workshopPresetIndex + 2);
          var presetRange = `${sheetName}!${presetCol}${i + 1}`;
          batchUpdate.push({
            range: presetRange,
            values: [[presetValue]],
          });
        }
        var enhancementName = row[workshopEnhancementNameCol];
        if (enhancementName.includes("Order in between")) {
          enhancementName = "Order in between";
        }
        if (enhancementName && enhancementName.startsWith("=")) {
          var parts = enhancementName.split(",");
          if (parts.length === 4) {
            enhancementName = parts[2].trim();
          } else {
            enhancementName = parts[parts.length - 1].trim();
          }
          enhancementName = enhancementName.replace(/['"()]/g, "");
        }
        if (enhancementName && oldWorkshopPlusRatios[enhancementName]) {
          var ratioData = oldWorkshopPlusRatios[enhancementName];
          ratiosToUpdate.push(ratioData);
        } else {
          ratiosToUpdate.push(["", "", "", "", "", "", "", "", "", ""]);
        }
        if (enhancementName && enhancementName === "Order in between") {
          break;
        }
      }
      if (ratiosToUpdate.length > 0) {
        var startRow = 2;
        var startCol = sheetRefs.columnToLetter(workshopEnhancementNameCol + 2);
        var endCol = sheetRefs.columnToLetter(workshopEnhancementNameCol + 11);
        var ratiosRange = `${sheetName}!${startCol}${startRow}:${endCol}${
          startRow + ratiosToUpdate.length - 1
        }`;
        batchUpdate.push({
          range: ratiosRange,
          values: ratiosToUpdate,
        });
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Workshop Plus Ratios updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for Workshop Plus Ratios`,
      };
    } catch (error) {
      var errorReport = errors.report("workshopWriter.updateWorkshopPlusRatios", error, {
        sheetName: sheetName,
        oldWorkshopPlusRatios: oldWorkshopPlusRatios,
        desiredRatiosData: desiredRatiosData,
      });
      return errors.fail(errorReport);
    }
  },
};
