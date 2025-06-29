const themes = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newThemeSpreadsheet) {
    importThemesData(sheetType, newThemeSpreadsheet)
    function importThemesData(sheetType, newThemeSpreadsheet) {
      var idType = sheetType + " ID"
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner"
      ]
      // Newly copied Theme sheet
      var newThemeIdSheet = newThemeSpreadsheet.getSheetByName("IDS")
      var newThemestats = newThemeSpreadsheet.getSheetByName("STATS")
      var newThemeVersion = newThemestats.getRange("A1").getValue()

      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newThemeIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldThemespreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldThemestats = oldThemespreadsheet.getSheetByName("STATS")
      var oldThemeVersion = oldThemestats.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldThemeVersion, newThemeVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        var oldThemes = oldThemespreadsheet.getSheetByName("Themes & Songs").getDataRange().getValues()
        var oldThemesNames = getOldUnlockedThemesNames(targetThemes, oldThemes)
        var newThemesSheet = newThemeSpreadsheet.getSheetByName("Themes & Songs")
        updateThemes(targetThemes, newThemesSheet, oldThemesNames)
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateThemes(targetThemes, newThemesSheet, oldThemesNames) {
      // Grab the old themes & songs that are unlocked
      var newThemes = newThemesSheet.getDataRange().getValues()

      // For each header, store {col, startRow} for quick reference
      var headerLocations = {}

      // Pre-scan to find header columns and their start rows
      for (var i = 0; i < newThemes.length; i++) {
        for (var j = 0; j < newThemes[i].length; j++) {
          var newThemeUnlocked = String(newThemes[i][j]).trim()
          if (targetThemes.indexOf(newThemeUnlocked) !== -1) {
            // If not already recorded for this col, store its location
            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = []
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 }) // +1 to start below header
          }
        }
      }

      // For each header, possibly in multiple places
      for (var key in headerLocations) {
        headerLocations[key].forEach(function(loc) {
          var checkboxCol = loc.col
          var startRow = loc.startRow
          var checkedSet = new Set((oldThemesNames[key] || []).map(String))
          var checkboxArr = []

          for (var row = startRow; row < newThemes.length; row++) {
            var newThemeName = newThemes[row][checkboxCol + 1]
            if (
              newThemeName === "" ||
              newThemeName === null ||
              typeof newThemeName === "undefined" ||
              targetThemes.indexOf(String(newThemeName).trim()) !== -1
            ) {
              break
            }
            checkboxArr.push([checkedSet.has(String(newThemeName))])
          }
          if (checkboxArr.length > 0) {
            newThemesSheet
              .getRange(startRow + 1, checkboxCol + 1, checkboxArr.length, 1)
              .setValues(checkboxArr)
          }
        })
      }
    }

    function getOldUnlockedThemesNames(targetThemes, oldThemes) {
      var oldThemesNames = {}
      targetThemes.forEach(function(header) {
        oldThemesNames[header] = []
      })
      var currentHeader = null
      var headerCol = -1

      // Loop through each column first, then rows
      for (var col = 0; col < oldThemes[0].length; col++) {
        for (var row = 0; row < oldThemes.length; row++) {
          var oldThemeUnlocked = oldThemes[row][col]
          // If cell is a header
          if (targetThemes.indexOf(String(oldThemeUnlocked).trim()) !== -1) {
            currentHeader = String(oldThemeUnlocked).trim()
            headerCol = col
            continue
          }
          // If we are collecting for a header, and current oldThemeUnlocked is true
          if (currentHeader && col === headerCol && oldThemeUnlocked) {
            var oldThemeName = oldThemes[row][col+1]
            oldThemesNames[currentHeader].push(oldThemeName)
          }
        }
      }

      return oldThemesNames
    }
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}