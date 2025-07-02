const guardians = {
  convertVersionFunctions: {},

  importData: function (newSheetID, oldSheetID) {
    function importGuardiansData(newSheetID, oldSheetID) {
      try {
        var sheetType = "Guardians";
        var idType = sheetType + " ID";
        var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];

        // Get new guardian version using SheetsAPI
        var newGuardianVersion = SheetsAPI.getValue(newSheetID, "EXPORT!A1");
        if (!newGuardianVersion) {
        console.log("Error getting new guardian version");
        return {
          success: false,
          message: "Error getting new guardian version"
        };
        }

        // Get ID Master spreadsheet info
        var idMasterInfo = shared.findSheetTypeID(
          newSheetID,
          "IDS"
        );
        if (!idMasterInfo || !idMasterInfo.id) {
          console.log("Could not find ID Master spreadsheet info");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet info"
          };
        }
        var idMasterID = shared.extractSheetId(
          idMasterInfo.id
        );
        if (!idMasterID) {
          console.log("Could not find ID Master spreadsheet");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet"
          };
        }
        // Get old guardian version using SheetsAPI
        var oldGuardianVersion = SheetsAPI.getValue(oldSheetID, "EXPORT!A1");
        if (!oldGuardianVersion) {
        console.log("Error getting old guardian version");
        return {
          success: false,
          message: "Error getting old guardian version"
        };
        }

        var versionCheck = shared.compareVersions(
          oldGuardianVersion,
          newGuardianVersion
        );
        if (versionCheck === 0) {
          console.log("Same Version");

          // Get header row from _IDS sheet using SheetsAPI
          var headerRowData = SheetsAPI.getValues(idMasterID, "_IDS!1:1")[0] || [];
          var importGuardianColStart = headerRowData.indexOf("Guardians");
          if (importGuardianColStart === -1) {
          console.log("Guardians column not found in header");
          return {
            success: false,
            message: "Guardians column not found in header"
          };
          }

          // Get old guardian levels data using SheetsAPI
          var colStart = shared.columnToLetter(importGuardianColStart + 1);
          var colEnd = shared.columnToLetter(importGuardianColStart + 5);
          var oldGuardianLevelsData = SheetsAPI.getValues(
            idMasterID,
            "_IDS!" + colStart + "2:" + colEnd
          );
          // Filter out empty rows
          var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
            row.some(
              (cell) =>
                cell !== null && cell !== undefined && String(cell || '').trim() !== ""
            )
          );

          var oldGuardians = getOldGuardians(targetGuardians, oldGuardianLevels);
          return updateGuardianLevels(
            targetGuardians,
            newSheetID,
            "Master Sheet",
            oldGuardians
          );
        }
        // else {// Else do something to convert old version to new one (Future me problem)
        // }
      } catch (error) {
        console.log("Error in importGuardiansData: " + error.toString());
        return {
          success: false,
          message: "Error in importGuardiansData: " + error.message,
        };
      }
    }

    function updateGuardianLevels(
      targetGuardians,
      newSheetID,
      sheetName,
      oldGuardians
    ) {
      // Get all data from Master Sheet using SheetsAPI
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData || sheetData.length < 2) return; // Need at least header and one row
      if (!sheetData || sheetData.length < 2) {
        return {
          success: false,
          message: "Not enough data in Master Sheet"
        };
      }

      var headerRow = sheetData[0];
      var guardianCol = headerRow.indexOf("Guardians");
      if (guardianCol === -1) {
      console.log("Guardians column not found in Master Sheet");
      return {
        success: false,
        message: "Guardians column not found in Master Sheet"
      };
      }

      // Get guardian data starting from row 2
      var guardianDataRange = [];
      for (var i = 1; i < sheetData.length; i++) {
        var row = sheetData[i];
        if (row.length > guardianCol + 4) {
          guardianDataRange.push([
            row[guardianCol + 1] || "",
            row[guardianCol + 2] || "",
            row[guardianCol + 3] || "",
            row[guardianCol + 4] || "",
            row[guardianCol + 5] || "",
          ]);
        }
      }

      // Filter out empty rows
      var newGuardianData = guardianDataRange.filter((row) =>
        row.some(
          (cell) =>
            cell !== null && cell !== undefined && String(cell || '').trim() !== ""
        )
      );

      var newGuardianUnlocked = [];
      var newGuardianLevel = [];
      
      for (var row = 0; row < newGuardianData.length; row++) {
        var rowData = newGuardianData[row];
        if (oldGuardians.hasOwnProperty(rowData[0])) {
          var oldGuardian = oldGuardians[rowData[0]];
          newGuardianUnlocked.push(
            [rowData[0]],
            [""],
            [oldGuardian.unlocked]
          );
          newGuardianUnlocked.push([rowData[0]]);
          newGuardianUnlocked.push([""]);
          newGuardianUnlocked.push([oldGuardian.unlocked]);
          for (var nextRow = row; nextRow < newGuardianData.length; nextRow++) {
            var nextRowData = newGuardianData[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var newGuardianProp = nextRowData[2];
            if (oldGuardian.props.hasOwnProperty(newGuardianProp)) {
              newGuardianLevel.push([oldGuardian.props[newGuardianProp]]);
            } else {
              newGuardianLevel.push([nextRowData[4]]);
            }
            if (nextRow == newGuardianData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newGuardianUnlocked.push([rowData[0]]);
        }
      }

      var batchUpdate = [];
      // Update the data using SheetsAPI
      if (newGuardianUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(guardianCol + 2);
        var unlockedRange =
          sheetName +
          "!" +
          unlockedCol +
          "2:" +
          unlockedCol +
          (1 + newGuardianUnlocked.length);
        batchUpdate.push({
          range: unlockedRange,
          values: newGuardianUnlocked,
        });
      }

      if (newGuardianLevel.length > 0) {
        var levelCol = shared.columnToLetter(guardianCol + 6);
        var levelRange =
          sheetName +
          "!" +
          levelCol +
          "2:" +
          levelCol +
          (1 + newGuardianLevel.length);
        batchUpdate.push({
          range: levelRange,
          values: newGuardianLevel,
        });
      }
      if (batchUpdate.length !== 0) {
        SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        console.log("Guardians updated successfully");
        return {
          success: true,
          message: "Guardians updated successfully",
        };
      }
      console.log("No updates needed for guardians");
      return {
        success: true,
        message: "No updates needed for guardians",
      };
    }

    function getOldGuardians(targetGuardians, oldGuardianLevels) {
      var guardians = {};
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked = oldGuardianLevels[row + 2][0];
          var guardian = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.props[key] = value;
            }
          }
          guardians[guardianName] = guardian;
        }
      }
      return guardians;
    }
    return importGuardiansData(newSheetID, oldSheetID);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
