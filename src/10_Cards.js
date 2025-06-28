const cards = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newCardsSpreadsheet) {
    importCardsData(sheetType, newCardsSpreadsheet)
    function importCardsData(sheetType, newCardsSpreadsheet) {
      var idType = sheetType + " ID"
      // Newly copied Cards sheet
      var newCardsIdSheet = newCardsSpreadsheet.getSheetByName("IDS")
      var newCardsExport = newCardsSpreadsheet.getSheetByName("EXPORT")
      var newCardsVersion = newCardsExport.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newCardsIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldCardsSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldCardsExport = oldCardsSpreadsheet.getSheetByName("EXPORT")
      var oldCardsVersion = oldCardsExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldCardsVersion, newCardsVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newCardsImportSheet = newCardsSpreadsheet.getSheetByName("_IDS")
        var headerRow = newCardsImportSheet.getRange(1, 1, 1, newCardsImportSheet.getLastColumn()).getValues()[0]
        var importCardsColStart = headerRow.indexOf("Cards") + 1
        var oldCardsLevels = newCardsImportSheet.getRange(2, importCardsColStart, newCardsImportSheet.getLastRow(), 3).getValues()
        oldCardsLevels = oldCardsLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var oldCardSlots = oldCardsExport.getRange("C2").getValue()

        var importCardsPresetsColStart = headerRow.indexOf("Cards Presets") + 1
        var oldCardsPresets = newCardsImportSheet.getRange(2, importCardsPresetsColStart, newCardsImportSheet.getLastRow(), 5).getValues()
        oldCardsPresets = oldCardsPresets.filter(row => row.some(cell => String(cell).trim() !== ""))
        var newCardsMasterSheet = newCardsSpreadsheet.getSheetByName("Master Sheet")
        var newcardsPresetSheet = newCardsSpreadsheet.getSheetByName("Card Preset")

        updateCardsLevels(newCardsMasterSheet, oldCardsLevels, oldCardSlots)
        updateCardsPresets(newcardsPresetSheet, oldCardsPresets)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateCardsLevels(newCardsMasterSheet, oldCardsLevels, oldCardSlots) {
      var lastRow = newCardsMasterSheet.getLastRow()
      if (lastRow < 2) return
      
      var headerRow = newCardsMasterSheet.getRange(1, 1, 1, newCardsMasterSheet.getLastColumn()).getValues()[0]
      var newCardNameCol = headerRow.indexOf("Card Name") + 1
      var numRows = lastRow - 1
      var newCardsLevels = newCardsMasterSheet.getRange(2, newCardNameCol, numRows, 3).getValues()
      var oldCards = {}
      oldCardsLevels.forEach(function(row) {
        var oldCardName = row[0]
        var oldLevel = row[1]
        var oldMastery = row[2]
        if (oldCardName) {
          oldCards[oldCardName] = [oldLevel, oldMastery]
        }
      })

      var newCards = []
      newCardsLevels.forEach(function(row) {
        var newCardName = row[0]
        if (newCardName === "Card Slot (Gems)") {
          newCards.push([oldCardSlots, ""])
        }
        else if (oldCards.hasOwnProperty(newCardName)) {
          newCards.push(oldCards[newCardName])
        }
        else {
          newCards.push([row[1], row[2]])
        }
      })
      newCardsMasterSheet.getRange(2, newCardNameCol + 1, newCards.length, newCards[0].length).setValues(newCards)
    }

    function updateCardsPresets(newcardsPresetSheet, oldCardsPresets) {
      var lastRow = newcardsPresetSheet.getLastRow()
      if (lastRow < 2) return
      var headerRow = newcardsPresetSheet.getRange(2, 1, 1, newcardsPresetSheet.getLastColumn()).getValues()[0]
      var newCardPresetNameIdxs = headerRow.map(function(cell, idx) {
        return String(cell).trim() !== "" ? idx : -1}).filter(function(idx) {
          return idx !== -1})
      var oldCardsPresetsHeaders = oldCardsPresets[0]
      var oldCardsPresetsCards = oldCardsPresets.slice(1)
      oldCardsPresetsHeaders.forEach(function(header, headerIdx) {
        var colIdx = newCardPresetNameIdxs[headerIdx] + 1
        newcardsPresetSheet.getRange(2, colIdx).setValue(header)
        var newCardsPresetsCards = oldCardsPresetsCards.map(function(row) {
          return [row[headerIdx]]
        })
        newcardsPresetSheet.getRange(3, colIdx + 1, newCardsPresetsCards.length, 1).setValues(newCardsPresetsCards)
      })
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}