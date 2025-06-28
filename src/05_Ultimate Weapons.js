const ultimate = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newUltimateSpreadsheet) {
    importUltimateData(sheetType, newUltimateSpreadsheet)
    function importUltimateData(sheetType, newUltimateSpreadsheet) {
      var idType = sheetType + " ID"
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ]
      // Newly copied Ultimate Weapons sheet
      var newUltimateIdSheet = newUltimateSpreadsheet.getSheetByName("IDS")
      var newUltimateExport = newUltimateSpreadsheet.getSheetByName("EXPORT")
      var newUltimateVersion = newUltimateExport.getRange("A1").getValue()
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newUltimateIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldUltimateSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldUltimateExport = oldUltimateSpreadsheet.getSheetByName("EXPORT")
      var oldUltimateVersion = oldUltimateExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldUltimateVersion, newUltimateVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var newUltimateImportSheet = newUltimateSpreadsheet.getSheetByName("_IDS")
        var headerRow = newUltimateImportSheet.getRange(1, 1, 1, newUltimateImportSheet.getLastColumn()).getValues()[0]
        var importUltimateColStart = headerRow.indexOf("UWs") + 1
        var oldUltimateLevels = newUltimateImportSheet.getRange(2, importUltimateColStart, newUltimateImportSheet.getLastRow(), 5).getValues()
        oldUltimateLevels = oldUltimateLevels.filter(row => row.some(cell => String(cell).trim() !== ""))
        var oldUltimate = getOldUltimateWeapons(targetWeapons, oldUltimateLevels)
        var newUltimateMasterSheet = newUltimateSpreadsheet.getSheetByName("Master Sheet")
        updateUltimateLevels(targetWeapons, newUltimateMasterSheet, oldUltimate)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateUltimateLevels(targetWeapons, newUltimateMasterSheet, oldUltimate) {
      var lastRow = newUltimateMasterSheet.getLastRow()
      if (lastRow < 2) return

      var headerRow = newUltimateMasterSheet.getRange(1, 1, 1, newUltimateMasterSheet.getLastColumn()).getValues()[0]
      var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1

      var newUltimateData = newUltimateMasterSheet.getRange(2, ultimateCol + 1, newUltimateMasterSheet.getLastRow(), 5).getValues()
      newUltimateData = newUltimateData.filter(row => row.some(cell => String(cell).trim() !== ""))
      var newUltimateUnlocked = []
      var newUltimateLevel = []
      for (var row = 0; row < newUltimateData.length; row++) {
        var rowData = newUltimateData[row]
        if (oldUltimate.hasOwnProperty(rowData[0])) {
          var oldWeapon = oldUltimate[rowData[0]]
          newUltimateUnlocked.push([rowData[0]])
          newUltimateUnlocked.push([""])
          newUltimateUnlocked.push([oldWeapon.unlocked])
          for (var nextRow = row; nextRow < newUltimateData.length; nextRow++) {
            var nextRowData = newUltimateData[nextRow]
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 2
              break
            }
            var newWeaponProp = nextRowData[2]
            if (oldWeapon.props.hasOwnProperty(newWeaponProp)) {
              newUltimateLevel.push([oldWeapon.props[newWeaponProp]])
            }
            else {
              newUltimateLevel.push([nextRowData[4]])
            }
            if (nextRow == newUltimateData.length - 1) {
              row = nextRow
            }
          }
        }
        else {
          newUltimateUnlocked.push([rowData[0]])
          // newUltimateLevel.push([rowData[4]])
        }
      }
      
      newUltimateMasterSheet.getRange(2, ultimateCol + 1, newUltimateUnlocked.length, 1).setValues(newUltimateUnlocked)
      newUltimateMasterSheet.getRange(2, ultimateCol + 5, newUltimateLevel.length, 1).setValues(newUltimateLevel)
    }

    function getOldUltimateWeapons(targetWeapons, oldUltimateLevels) {
      var weapons = {}
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0]
        // Only proceed if weaponName is in targetWeapons
        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0]
          var weapon = {
            unlocked: unlocked,
            props: {}
          }

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow]
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1
              break
            }
            var key = nextRowData[2]
            var value = nextRowData[4]
            if (key && value) {
              weapon.props[key] = value
            }
          }
          weapons[weaponName] = weapon
        }
      }
      return weapons
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}