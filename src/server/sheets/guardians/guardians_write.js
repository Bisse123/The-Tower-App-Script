const guardiansWriter = {

  /**
   * Builds the batch update that writes GuardianLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldGuardians
   * @param {Object} masterSheetData
   * @param {Object} dvtNamedRangesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateGuardianLevels: function (
    sheetName,
    oldGuardians,
    masterSheetData,
    dvtNamedRangesData
  ) {
    try {
      console.log("Called: guardiansWriter.updateGuardianLevels");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      if (!masterSheetData || masterSheetData.length < 2) {
        return {
          success: false,
          message: "Not enough data in Master Sheet™",
        };
      }

      var headerRow = masterSheetData[0];
      var guardianCol = headerRow.indexOf("Chips") + 1;

      if (guardianCol === 0) {
        console.log(`Guardian Weapon column not found`);
        return {
          success: false,
          message: `Guardian Weapon column not found`,
        };
      }

      var startCol = guardianCol + 1;
      var endCol = startCol;

      var presetColumnMapping = [];
      var firstPresetIndex = startCol + 2;

      if (firstPresetIndex === -1) {
        console.log(`Preset columns not found in Master Sheet`);
        return {
          success: false,
          message: `Preset columns not found in Master Sheet™`,
        };
      }

      var batchUpdate = [];

      var presetSlots = [];
      headerRow.forEach(function (header, index) {
        if (index < firstPresetIndex || !String(header).trim() || presetSlots.length >= oldGuardians.presetNames.length) {
          return;
        }
        presetSlots.push({ header: String(header).trim(), colIndex: index });
      });

      oldGuardians.presetNames.forEach(function (presetName, slot) {
        var presetSlot = presetSlots[slot];
        if (!presetName || !presetSlot) {
          return;
        }
        var colIndex = presetSlot.colIndex;
        endCol = colIndex + 1;
        presetColumnMapping.push({
          presetName: presetName,
          equippedColIndex: colIndex,
          levelColIndex: endCol,
        });
        if (presetSlot.header !== presetName) {
          batchUpdate.push({
            range: `${sheetName}!${sheetRefs.columnToLetter(colIndex + 1)}1`,
            values: [[presetName]],
          });
        }
      });

      var newGuardianData = masterSheetData
        .slice(1)
        .map(function (row) {
          return row.slice(startCol - 1, endCol);
        })
        .filter(function (row) {
          return row.some(function (cell) {
            return (
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
            );
          });
        });

      if (!newGuardianData || newGuardianData.length === 0) {
        return {
          success: false,
          message: "Could not read guardian data",
        };
      }

      var newGuardianUnlocked = [];
      var newGuardianLevels = {};
      var newGuardianEquipped = {};

      presetColumnMapping.forEach(function (presetMap) {
        newGuardianLevels[presetMap.presetName] = [];
        newGuardianEquipped[presetMap.presetName] = [];
      });

      var currentGuardianName = null;
      var currentGuardian = null;
      var guardianRow = -1;

      for (var row = 0; row < newGuardianData.length; row++) {
        var rowData = newGuardianData[row];
        var rowGuardianName = String(rowData[0] || "").trim();

        if (rowGuardianName && targetGuardians.includes(rowGuardianName)) {
          currentGuardianName = rowGuardianName;
          currentGuardian = oldGuardians.data.hasOwnProperty(rowGuardianName)
            ? oldGuardians.data[rowGuardianName]
            : null;
          guardianRow = row;
        }

        if (currentGuardian && row === guardianRow) {
          newGuardianUnlocked.push([currentGuardianName]);
        } else if (currentGuardian && row === guardianRow + 2) {
          newGuardianUnlocked.push([currentGuardian.unlocked]);
        } else {
          newGuardianUnlocked.push([null]);
        }

        var newGuardianProp = rowData[2];
        var dvtGuardianRanges = currentGuardianName
          ? dvtNamedRangesData[currentGuardianName]
          : null;

        presetColumnMapping.forEach(function (presetMap) {
          var presetName = presetMap.presetName;
          var presetData =
            currentGuardian &&
            currentGuardian.presets &&
            currentGuardian.presets[presetName]
              ? currentGuardian.presets[presetName]
              : null;

          if (
            presetData &&
            row === guardianRow &&
            presetData.hasOwnProperty("equipped")
          ) {
            newGuardianEquipped[presetName].push([presetData.equipped]);
          } else {
            newGuardianEquipped[presetName].push([null]);
          }

          if (
            !presetData ||
            !presetData.hasOwnProperty("props") ||
            !newGuardianProp ||
            !presetData.props.hasOwnProperty(newGuardianProp) ||
            !dvtGuardianRanges
          ) {
            newGuardianLevels[presetName].push([null]);
            return;
          }

          var dvtPropValue = dropdownUtils.getDVTValue(
            presetData.props[newGuardianProp],
            dvtGuardianRanges[newGuardianProp]
          );
          newGuardianLevels[presetName].push([dvtPropValue]);
        });
      }

      if (newGuardianUnlocked.length > 0) {
        var unlockedCol = sheetRefs.columnToLetter(guardianCol + 1);
        batchUpdate.push({
          range: `${sheetName}!${unlockedCol}2:${unlockedCol}${
            newGuardianUnlocked.length + 1
          }`,
          values: newGuardianUnlocked,
        });
      }

      presetColumnMapping.forEach(function (presetMap) {
        var levels = newGuardianLevels[presetMap.presetName];
        if (!levels || levels.length === 0) {
          return;
        }
        var levelCol = sheetRefs.columnToLetter(presetMap.levelColIndex + 1);
        batchUpdate.push({
          range: `${sheetName}!${levelCol}2:${levelCol}${levels.length + 1}`,
          values: levels,
        });
      });

      presetColumnMapping.forEach(function (presetMap) {
        var equipped = newGuardianEquipped[presetMap.presetName];
        if (
          !equipped ||
          !equipped.some(function (value) {
            return value[0] !== null && value[0] !== undefined;
          })
        ) {
          return;
        }
        var equippedCol = sheetRefs.columnToLetter(presetMap.equippedColIndex + 1);
        batchUpdate.push({
          range: `${sheetName}!${equippedCol}2:${equippedCol}${
            equipped.length + 1
          }`,
          values: equipped,
        });
      });

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `Guardians updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for guardians`,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.updateGuardianLevels", error, {
        sheetName: sheetName,
        oldGuardians: oldGuardians,
        masterSheetData: masterSheetData,
        dvtNamedRangesData: dvtNamedRangesData,
      });
      return errors.fail(errorReport);
    }
  },
};
