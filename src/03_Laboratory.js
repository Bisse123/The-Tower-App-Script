const lab = {
  convertVersionFunctions: {
    },

  importData: function importData(sheetType, newLabSpreadsheet) {
    importLabData(sheetType, newLabSpreadsheet)
    function importLabData(sheetType, newLabSpreadsheet) {
      var idType = sheetType + " ID"
      // Newly copied lab sheet
      var newLabIdSheet = newLabSpreadsheet.getSheetByName("IDS")
      var newLabExport = newLabSpreadsheet.getSheetByName("EXPORT")
      var newLabVersion = newLabExport.getRange("A1").getValue()

      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newLabIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")

      var oldLabSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldLabExport = oldLabSpreadsheet.getSheetByName("EXPORT")
      var oldLabVersion = oldLabExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldLabVersion, newLabVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newLabImportSheet = newLabSpreadsheet.getSheetByName("_IDS")
        var headerRow = newLabImportSheet.getRange(1, 1, 1, newLabImportSheet.getLastColumn()).getValues()[0]
        var importLabColStart = headerRow.indexOf("Labs") + 1
        var oldLabLevels = newLabImportSheet.getRange(2, importLabColStart, newLabImportSheet.getLastRow(), 3).getValues()
        oldLabLevels = oldLabLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var newLabMasterSheet = newLabSpreadsheet.getSheetByName("Master Sheet")
        updateLabLevels(newLabMasterSheet, oldLabLevels)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateLabLevels(labMasterSheet, labUpdates) {
      var headerValues = ["Labs"]
      var lastRow = labMasterSheet.getLastRow()
      if (lastRow < 2) return

      // Dynamically find columns where header is in headerValues
      var headerRow = labMasterSheet.getRange(1, 1, 1, labMasterSheet.getLastColumn()).getValues()[0]
      var columnsToCheck = []
      for (var i = 0; i < headerRow.length; i++) {
        if (headerValues.includes(headerRow[i])) {
          columnsToCheck.push(i + 1)
        }
      }

      // Ignores the very last row with sums (Should never reach it but done for safety anyways)
      var numRows = lastRow - 2
      var colMax = Math.max(...columnsToCheck) + 2
      var values = labMasterSheet.getRange(2, 1, numRows, colMax).getValues()

      // Prepare update map from labName to update values
      var updateMap = {}
      labUpdates.forEach(function(update) {
        updateMap[update[0]] = [update[1], update[2]]
      })

      // Iterate each "Labs" column
      columnsToCheck.forEach(function(col) {
        var updates = []
        // Find labNames in each column
        for (var row = 0; row < values.length; row++) {
          var cellValue = values[row][col - 1]
          if (cellValue === "") break
          var update = updateMap[cellValue]
          // If labName is found then push update regardless of the imported levels
          if (update) {
            updates.push([update[0], update[1]])
          }
          else {
            updates.push([values[row][col], values[row][col + 1]])
          }
        }

        // Write the imported values to the cells in column "Level" and "Target"
        if (updates.length) {
          labMasterSheet
            .getRange(2, col + 1, updates.length, 2)
            .setValues(updates)
        }
      })
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}