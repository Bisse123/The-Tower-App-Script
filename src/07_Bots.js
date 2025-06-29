const bots = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newBotSpreadsheet) {
    importBotsData(sheetType, newBotSpreadsheet)
    function importBotsData(sheetType, newBotSpreadsheet) {
      var idType = sheetType + " ID"
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
      ]
      // Newly copied Bots sheet
      var newBotIdSheet = newBotSpreadsheet.getSheetByName("IDS")
      var newBotExport = newBotSpreadsheet.getSheetByName("EXPORT")
      var newBotVersion = newBotExport.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newBotIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldBotSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldBotExport = oldBotSpreadsheet.getSheetByName("EXPORT")
      var oldBotVersion = oldBotExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldBotVersion, newBotVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newBotImportSheet = newBotSpreadsheet.getSheetByName("_IDS")
        var headerRow = newBotImportSheet.getRange(1, 1, 1, newBotImportSheet.getLastColumn()).getValues()[0]
        var importbotColStart = headerRow.indexOf("Bots") + 1
        var oldBotLevels = newBotImportSheet.getRange(2, importbotColStart, newBotImportSheet.getLastRow(), 5).getValues()
        oldBotLevels = oldBotLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var oldBots = getOldBots(targetBots, oldBotLevels)
        var newBotMasterSheet = newBotSpreadsheet.getSheetByName("Master Sheet")
        updateBotLevels(targetBots, newBotMasterSheet, oldBots)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateBotLevels(targetBots, newBotMasterSheet, oldBots) {
      var lastRow = newBotMasterSheet.getLastRow()
      if (lastRow < 2) return

      var headerRow = newBotMasterSheet.getRange(1, 1, 1, newBotMasterSheet.getLastColumn()).getValues()[0]
      var botCol = headerRow.indexOf("Bot") + 1

      var newBotData = newBotMasterSheet.getRange(2, botCol + 1, newBotMasterSheet.getLastRow(), 5).getValues()
      newBotData = newBotData.filter(row => row.some(cell => String(cell).trim() !== ""))
      var newBotUnlocked = []
      var newBotLevel = []
      for (var row = 0; row < newBotData.length; row++) {
        var rowData = newBotData[row]
        if (oldBots.hasOwnProperty(rowData[0])) {
          var oldWeapon = oldBots[rowData[0]]
          newBotUnlocked.push([rowData[0]])
          newBotUnlocked.push([""])
          newBotUnlocked.push([""])
          newBotUnlocked.push([oldWeapon.unlocked])
          for (var nextRow = row; nextRow < newBotData.length; nextRow++) {
            var nextRowData = newBotData[nextRow]
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1
              break
            }
            var newWeaponProp = nextRowData[2]
            if (oldWeapon.props.hasOwnProperty(newWeaponProp)) {
              newBotLevel.push([oldWeapon.props[newWeaponProp]])
            }
            else {
              newBotLevel.push([nextRowData[4]])
            }
            if (nextRow == newBotData.length - 1) {
              row = nextRow
            }
          }
        }
        else {
          newBotUnlocked.push([rowData[0]])
          // newBotLevel.push([rowData[4]])
        }
      }
      
      newBotMasterSheet.getRange(2, botCol + 1, newBotUnlocked.length, 1).setValues(newBotUnlocked)
      newBotMasterSheet.getRange(2, botCol + 5, newBotLevel.length, 1).setValues(newBotLevel)
    }

    function getOldBots(targetBots, oldBotLevels) {
      var bots = {}
      for (var row = 0; row < oldBotLevels.length; row++) {
        var weaponName = oldBotLevels[row][0]
        // Only proceed if weaponName is in targetBots
        if (weaponName && targetBots.includes(weaponName)) {
          var unlocked = oldBotLevels[row + 3][0]
          var weapon = {
            unlocked: unlocked,
            props: {}
          }

          for (nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
            var nextRowData = oldBotLevels[nextRow]
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1
              break
            }
            var key = nextRowData[2]
            var value = nextRowData[4]
            if (key && value) {
              weapon.props[key] = value
            }
          }
          bots[weaponName] = weapon
        }
      }
      return bots
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}