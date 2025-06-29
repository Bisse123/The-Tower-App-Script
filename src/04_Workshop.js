const workshop = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newWorkshopSpreadsheet) {
    importWorkshopData(sheetType, newWorkshopSpreadsheet)
    function importWorkshopData(sheetType, newWorkshopSpreadsheet) {
      var idType = sheetType + " ID"
      // Newly copied workshop sheet
      var newWorkshopIdSheet = newWorkshopSpreadsheet.getSheetByName("IDS")
      var newWorkshopExport = newWorkshopSpreadsheet.getSheetByName("EXPORT")
      var newWorkshopVersion = newWorkshopExport.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newWorkshopIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")

      var oldWorkshopSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldWorkshopExport = oldWorkshopSpreadsheet.getSheetByName("EXPORT")
      var oldWorkshopVersion = oldWorkshopExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldWorkshopVersion, newWorkshopVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newWorkshopImportSheet = newWorkshopSpreadsheet.getSheetByName("_IDS")
        var headerRow = newWorkshopImportSheet.getRange(1, 1, 1, newWorkshopImportSheet.getLastColumn()).getValues()[0]
        var importWorkshopColStart = headerRow.indexOf("WS") + 1
        var oldWorkshopLevels = newWorkshopImportSheet.getRange(2, importWorkshopColStart, newWorkshopImportSheet.getLastRow(), 4).getValues()
        oldWorkshopLevels = oldWorkshopLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        
        var importWorkshopPlusColStart = headerRow.indexOf("WS+") + 1
        var oldWorkshopPlusLevels = newWorkshopImportSheet.getRange(2, importWorkshopPlusColStart, newWorkshopImportSheet.getLastRow(), 3).getValues()
        oldWorkshopPlusLevels = oldWorkshopPlusLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var newWorkshopMasterSheet = newWorkshopSpreadsheet.getSheetByName("Master Sheet")
        updateWorkshopLevels(newWorkshopMasterSheet, oldWorkshopLevels, oldWorkshopPlusLevels)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateWorkshopLevels(workshopMasterSheet, workshopLevels, workshopPlusLevels) {
      var lastRow = workshopMasterSheet.getLastRow()
      if (lastRow < 2) return

      // Find header row and relevant columns
      var headerRow = workshopMasterSheet.getRange(1, 1, 1, workshopMasterSheet.getLastColumn()).getValues()[0]
      var upgradeCol = headerRow.indexOf("Workshop Upgrade") + 1
      var enhancementCol = headerRow.indexOf("Workshop Enhancement") + 1

      // Split workshopLevels into workshopUnlocked and workshopLevelsSplit
      // workshopUnlocked: first element of each sublist
      // workshopLevelsSplit: 3rd and 4th elements (indices 2 and 3) of each sublist
      var workshopUnlocked = workshopLevels.map(function(sublist) {
        return [sublist[0]]
      })
      var workshopLevelsSplit = workshopLevels.map(function(sublist) {
        return [sublist[2], sublist[3]]
      })
      var workshopPlusLevelsSplit = workshopPlusLevels.map(function(sublist) {
        return[sublist[2]]
      })
      // Write workshopUnlocked: column before "Workshop Upgrade"
      if (upgradeCol > 1 && workshopUnlocked.length) {
        workshopMasterSheet.getRange(2, upgradeCol - 1, workshopUnlocked.length, 1).setValues(workshopUnlocked)
      }

      // Write workshopLevelsSplit: 1 columns after "Workshop Upgrade"
      if (upgradeCol > 0 && workshopLevelsSplit.length) {
        workshopMasterSheet.getRange(2, upgradeCol + 1, workshopLevelsSplit.length, workshopLevelsSplit[0].length).setValues(workshopLevelsSplit)
      }

      // Write workshopPlusLevelsSplit: 2 columns after "Workshop Enhancement"
      if (enhancementCol > 0 && workshopPlusLevelsSplit.length) {
        workshopMasterSheet.getRange(2, enhancementCol + 2, workshopPlusLevelsSplit.length, workshopPlusLevelsSplit[0].length).setValues(workshopPlusLevelsSplit)
      }
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}