const relics = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newRelicSpreadsheet) {
    importRelicsData(sheetType, newRelicSpreadsheet)
    function importRelicsData(sheetType, newRelicSpreadsheet) {
      var idType = sheetType + " ID"
      // Newly copied relic sheet
      var newRelicIdSheet = newRelicSpreadsheet.getSheetByName("IDS")
      var newRelicStats = newRelicSpreadsheet.getSheetByName("STATS")
      var newRelicVersion = newRelicStats.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newRelicIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldRelicSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldRelicStats = oldRelicSpreadsheet.getSheetByName("STATS")
      var oldRelicVersion = oldRelicStats.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldRelicVersion, newRelicVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var oldRelicsSheet = oldRelicSpreadsheet.getSheetByName("Relics")
        var oldRelicLastRow = oldRelicsSheet.getLastRow()
        var oldRelicLastCol = oldRelicsSheet.getLastColumn()
        var oldRelicHeaderRow = null
        var oldRelicNameCol = null
        var oldRelicUnlockedCol = null

        // Scan each row to find the header
        for (var row = 1; row <= oldRelicLastRow; row++) {
          var rowValues = oldRelicsSheet.getRange(row, 1, 1, oldRelicLastCol).getValues()[0]
          var relicNameIndex = rowValues.indexOf("Relic Name")
          var relicUnlockedIndex = rowValues.indexOf("Unlocked") + 1
          if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
            oldRelicHeaderRow = row
            oldRelicNameCol = relicNameIndex + 1
            oldRelicUnlockedCol = relicUnlockedIndex + 1
            break
          }
        }
        if (oldRelicHeaderRow) {
          var numRows = oldRelicLastRow - oldRelicHeaderRow
          var numCols = oldRelicUnlockedCol - oldRelicNameCol + 1
          var oldRelics = oldRelicsSheet.getRange(3, oldRelicNameCol, numRows, numCols).getValues()
          oldRelics = oldRelics.filter(row => row.some(cell => String(cell).trim() !== ""))
          var newRelicsSheet = newRelicSpreadsheet.getSheetByName("Relics")
          updateRelics(newRelicsSheet, oldRelics)
        }
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateRelics(newRelicsSheet, oldRelics) {
      var lastRow = newRelicsSheet.getLastRow()
      if (lastRow < 2) return

      // Dynamically find columns where header is in headerValues
      var headerRow = newRelicsSheet.getRange(2, 1, 1, newRelicsSheet.getLastColumn()).getValues()[0]
      var newRelicNameCol = headerRow.indexOf("Relic Name") + 1
      var newRelicUnlockedCol = headerRow.indexOf("Unlocked") + 1
      
      oldRelicsNames = []
      oldRelics.forEach(function(relic) {
        if (relic[3]) {
          oldRelicsNames.push(relic[0])
        }
      })

      var newRelicsNames = newRelicsSheet.getRange(3, newRelicNameCol, newRelicsSheet.getLastRow(), 1).getValues()
      newRelicsNames = newRelicsNames.filter(row => row.some(cell => String(cell).trim() !== ""))
      var newRelicsUnlocked = []
      newRelicsNames.forEach(function(relic) {
        if (oldRelicsNames.includes(relic[0])) {
          newRelicsUnlocked.push([true])
        }
        else {
          newRelicsUnlocked.push([false])
        }
      })
      Logger.log(newRelicsUnlocked)
      newRelicsSheet.getRange(3, newRelicUnlockedCol, newRelicsNames.length, 1).setValues(newRelicsUnlocked)
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}