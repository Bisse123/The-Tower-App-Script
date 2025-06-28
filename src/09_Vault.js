const vault = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newVaultSpreadsheet) {
    importVaultData(sheetType, newVaultSpreadsheet)
    function importVaultData(sheetType, newVaultSpreadsheet) {
      var idType = sheetType + " ID"
      // Newly copied Vault sheet
      var newVaultIdSheet = newVaultSpreadsheet.getSheetByName("IDS")
      var newVaultStats = newVaultSpreadsheet.getSheetByName("STATS")
      var newVaultVersion = newVaultStats.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newVaultIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldVaultSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldVaultStats = oldVaultSpreadsheet.getSheetByName("STATS")
      var oldVaultVersion = oldVaultStats.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldVaultVersion, newVaultVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var oldVaultHarmonySheet = oldVaultSpreadsheet.getSheetByName("Harmony")
        var newVaultHarmonySheet = newVaultSpreadsheet.getSheetByName("Harmony")
        var vaultHarmonyHeaderPattern = ["U", "Value","Bonus Type"]
        var oldVaultHarmony = getOldVault(oldVaultHarmonySheet, vaultHarmonyHeaderPattern, 1)

        var oldVaultPowerSheet = oldVaultSpreadsheet.getSheetByName("Power")
        var newVaultPowerSheet = newVaultSpreadsheet.getSheetByName("Power")
        var vaultPowerHeaderPattern = ["U", "", "Value","Bonus Type"]
        var oldVaultPower = getOldVault(oldVaultPowerSheet, vaultPowerHeaderPattern, 1)

        updateVault(newVaultHarmonySheet, oldVaultHarmony, vaultHarmonyHeaderPattern, 1)
        updateVault(newVaultPowerSheet, oldVaultPower, vaultPowerHeaderPattern, 1)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateVault(newVaultSheet, oldVault, vaultPattern, skipRows) {
      var tierUnlock = ["Tier x2 Unlock", "Tier x3 Unlock"]
      var lastRow = newVaultSheet.getLastRow()
      if (lastRow < 2) return
      
      var newVaultValues = newVaultSheet.getDataRange().getValues()
      var newVaultHeaders = newVaultValues[skipRows]
      var newVaultData = newVaultValues.slice(skipRows + 1)

      var newHeaderIndices = findHeaderIndices(newVaultHeaders, vaultPattern)

      // Iterate through data (excluding headers)
      var newVault = {}
      for (var r = 0; r < newVaultData.length; r++) {
        var row = newVaultData[r]
        for (var t = 0; t < newHeaderIndices.length; t++) {
          var uIdx = newHeaderIndices[t][vaultPattern.indexOf("U")]
          var valueIdx = newHeaderIndices[t][vaultPattern.indexOf("Value")]
          var bonusTypeIdx = newHeaderIndices[t][vaultPattern.indexOf("Bonus Type")]
          var u = row[uIdx]
          var value = row[valueIdx]
          var bonusType = row[bonusTypeIdx]
          var key = bonusType || value
          if (oldVault.hasOwnProperty(key)) {
            u = oldVault[key].shift()
            if (oldVault[key].length === 0) {
              delete oldVault[key]
            }
          }
          if (tierUnlock.includes(key)) {
            newVaultSheet.getRange(r + skipRows + 2, uIdx + 1).setValue(u)
          }
          if (!newVault.hasOwnProperty(uIdx)) {
            newVault[uIdx] = []
          }
          newVault[uIdx].push([u])
        }
      }

      // Recalculate Formulas for datavalidation
      SpreadsheetApp.flush()

      Object.keys(newVault).forEach(function(colKey) {
        var colIdx = parseInt(colKey, 10) + 1
        var values = newVault[colKey]
        newVaultSheet.getRange(3, colIdx, values.length, 1).setValues(values)
      })
    }

    function getOldVault(oldVaultSheet, oldVaultPattern, skipRows) {
      var oldVaultValues = oldVaultSheet.getDataRange().getValues()
      var oldVaultHeaders = oldVaultValues[skipRows]
      var oldVaultData = oldVaultValues.slice(skipRows + 1)
      var oldVault = {}
      // Find the indices of all "U", "Value", and "Bonus Type" columns by header
      var oldHeaderIndices = findHeaderIndices(oldVaultHeaders, oldVaultPattern)

      // Iterate through data (excluding headers)
      for (var r = 0; r < oldVaultData.length; r++) {
        var row = oldVaultData[r]
        for (var t = 0; t < oldHeaderIndices.length; t++) {
          var uIdx = oldHeaderIndices[t][oldVaultPattern.indexOf("U")]
          var valueIdx = oldHeaderIndices[t][oldVaultPattern.indexOf("Value")]
          var bonusTypeIdx = oldHeaderIndices[t][oldVaultPattern.indexOf("Bonus Type")]
          var u = row[uIdx]
          var value = row[valueIdx]
          var bonusType = row[bonusTypeIdx]
          var key = bonusType || value
          if (key && isNaN(key)) {
            if (!oldVault.hasOwnProperty(key)) {
              oldVault[key] = []
            }
            oldVault[key].push(u)
          }
        }
      }
      return oldVault
    }

    function findHeaderIndices(headers, pattern) {
      var indices = [];
      for (var i = 0; i <= headers.length - pattern.length; i++) {
        var match = true;
        for (var j = 0; j < pattern.length; j++) {
          if (headers[i + j] !== pattern[j]) {
            match = false;
            break;
          }
        }
        if (match) {
          indices.push(Array.from({length: pattern.length}, (_, k) => i + k));
        }
      }
      return indices;
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}