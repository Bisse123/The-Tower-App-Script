const guardians = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newGuardianSpreadsheet) {
    importGuardiansData(sheetType, newGuardianSpreadsheet)
    function importGuardiansData(sheetType, newGuardianSpreadsheet) {
      var sheetType = "Guardians"
      var idType = sheetType + " ID"
      var targetGuardians = [
        "Attack",
        "Ally",
        "Steal",
        "Fetch",
      ]
      // Newly copied Guardian sheet
      var newGuardianIdSheet = newGuardianSpreadsheet.getSheetByName("IDS")
      var newGuardianExport = newGuardianSpreadsheet.getSheetByName("EXPORT")
      var newGuardianVersion = newGuardianExport.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newGuardianIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldGuardianSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldGuardianExport = oldGuardianSpreadsheet.getSheetByName("EXPORT")
      var oldGuardianVersion = oldGuardianExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldGuardianVersion, newGuardianVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newGuardianImportSheet = idMasterSpreadsheet.getSheetByName("_IDS")
        var headerRow = newGuardianImportSheet.getRange(1, 1, 1, newGuardianImportSheet.getLastColumn()).getValues()[0]
        var importGuardianColStart = headerRow.indexOf("Guardians") + 1
        var oldGuardianLevels = newGuardianImportSheet.getRange(2, importGuardianColStart, newGuardianImportSheet.getLastRow(), 5).getValues()
        oldGuardianLevels = oldGuardianLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var oldGuardians = getOldGuardians(targetGuardians, oldGuardianLevels)
        var newGuardianMasterSheet = newGuardianSpreadsheet.getSheetByName("Master Sheet")
        updateGuardianLevels(targetGuardians, newGuardianMasterSheet, oldGuardians)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateGuardianLevels(targetGuardians, newGuardianMasterSheet, oldGuardians) {
      var lastRow = newGuardianMasterSheet.getLastRow()
      if (lastRow < 2) return

      var headerRow = newGuardianMasterSheet.getRange(1, 1, 1, newGuardianMasterSheet.getLastColumn()).getValues()[0]
      var guardianCol = headerRow.indexOf("Guardians") + 1

      var newGuardianData = newGuardianMasterSheet.getRange(2, guardianCol + 1, newGuardianMasterSheet.getLastRow(), 5).getValues()
      newGuardianData = newGuardianData.filter(row => row.some(cell => String(cell).trim() !== ""))
      var newGuardianUnlocked = []
      var newGuardianLevel = []
      for (var row = 0; row < newGuardianData.length; row++) {
        var rowData = newGuardianData[row]
        if (oldGuardians.hasOwnProperty(rowData[0])) {
          var oldGuardian = oldGuardians[rowData[0]]
          newGuardianUnlocked.push([rowData[0]])
          newGuardianUnlocked.push([""])
          newGuardianUnlocked.push([oldGuardian.unlocked])
          for (var nextRow = row; nextRow < newGuardianData.length; nextRow++) {
            var nextRowData = newGuardianData[nextRow]
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1
              break
            }
            var newGuardianProp = nextRowData[2]
            if (oldGuardian.props.hasOwnProperty(newGuardianProp)) {
              newGuardianLevel.push([oldGuardian.props[newGuardianProp]])
            }
            else {
              newGuardianLevel.push([nextRowData[4]])
            }
            if (nextRow == newGuardianData.length - 1) {
              row = nextRow
            }
          }
        }
        else {
          newGuardianUnlocked.push([rowData[0]])
          // newGuardianLevel.push([rowData[4]])
        }
      }
      
      newGuardianMasterSheet.getRange(2, guardianCol + 1, newGuardianUnlocked.length, 1).setValues(newGuardianUnlocked)
      newGuardianMasterSheet.getRange(2, guardianCol + 5, newGuardianLevel.length, 1).setValues(newGuardianLevel)
    }

    function getOldGuardians(targetGuardians, oldGuardianLevels) {
      var guardians = {}
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0]
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked = oldGuardianLevels[row + 2][0]
          var guardian = {
            unlocked: unlocked,
            props: {}
          }

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow]
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1
              break
            }
            var key = nextRowData[2]
            var value = nextRowData[4]
            if (key && value) {
              guardian.props[key] = value
            }
          }
          guardians[guardianName] = guardian
        }
      }
      return guardians
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}