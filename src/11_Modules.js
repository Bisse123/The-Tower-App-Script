const modules = {
  convertVersionFunctions: {
      3: this.convertVersion3,
    },

  importData: function(sheetType, newModulesSpreadsheet) {
    importModulesData(sheetType, newModulesSpreadsheet)
    function importModulesData(sheetType, newModulesSpreadsheet) {
      var idType = sheetType + " ID"
      var targetModuleTypes = [
        "cannon",
        "armor",
        "generator",
        "core"
      ]
      // Newly copied Modules sheet
      var newModulesIdSheet = newModulesSpreadsheet.getSheetByName("IDS")
      var newModulesExport = newModulesSpreadsheet.getSheetByName("EXPORT")
      var newModulesVersion = newModulesExport.getRange("A1").getValue()
      var newModulesInventorySheet = newModulesSpreadsheet.getSheetByName("Modules Inventory")
      var newModulesPresetsSheet = newModulesSpreadsheet.getSheetByName("Modules Presets")
      
      // ID Master sheet
      var idMasterSpreadsheet = shared.openSpreadsheet(newModulesIdSheet)
      var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
      
      var oldModulesSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
      var oldModulesExport = oldModulesSpreadsheet.getSheetByName("EXPORT")
      var oldModulesVersion = oldModulesExport.getRange("A1").getValue()
      var versionCheck = shared.compareVersions(oldModulesVersion, newModulesVersion)
      if (versionCheck === 0) {
        Logger.log("Same Version")
        // Grabs Level and Target from _IDS sheet in the new sheet
        
        var oldModulesInventorySheet = oldModulesSpreadsheet.getSheetByName("Modules Inventory")
        var oldModulesInventoryValues = oldModulesInventorySheet.getDataRange().getValues()
        var oldModulesInventory = getOldModulesInventory(targetModuleTypes, oldModulesInventoryValues)
        updateModulesInventory(targetModuleTypes, newModulesInventorySheet, oldModulesInventory)

        var oldModulesPresetsSheet = oldModulesSpreadsheet.getSheetByName("Modules Presets")
        var oldModulesPresetsValues = oldModulesPresetsSheet.getDataRange().getValues()
        var oldModulesPresets = getOldModulesPresets(targetModuleTypes, oldModulesPresetsValues)
        updateModulesPresets(targetModuleTypes, newModulesPresetsSheet, oldModulesPresets)

      }
      // Else do something to convert old version to new one (Future me problem)
      else {
        var {oldModulesInventory, oldModulesPresets} = this.convertVersionFunctions[versionCheck](oldModulesSpreadsheet)
        updateModulesInventory(targetModuleTypes, newModulesInventorySheet, oldModulesInventory)
        updateModulesPresets(targetModuleTypes, newModulesPresetsSheet, oldModulesPresets)
      }
      // Check version to figure out which convert function to use
      // Potentially do something where if new version is v4 and old is v2
      // you can do convert v2 -> v3 and then convert v3 -> v4
      // Means more calculations but less rewriting of convert functions when new version are released
    }

    function updateModulesPresets(targetModuleTypes, newModulesPresetsSheet, oldModulesPresets) {
      var newModulePresetsValues = newModulesPresetsSheet.getDataRange().getValues()
      var newModuleTypeIndex = findModuleTypesRowIndex(targetModuleTypes, newModulePresetsValues)
      targetModuleTypes.forEach(function(moduleType) {
        if (oldModulesPresets.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType] + 1
          if (typeof rowIdx === "undefined") return
          var row = newModulePresetsValues[rowIdx]
          for (var col = 0; col < row.length; col++) {
            if (String(row[col]).trim() === "Module Name") {
              var presetName = String(newModulePresetsValues[rowIdx - 1][col]).trim()
              if (presetName && oldModulesPresets[moduleType].hasOwnProperty(presetName)) {
                newModulesPresetsSheet.getRange(rowIdx + 2, col + 1).setValue(oldModulesPresets[moduleType][presetName])
              }
            }
          }
        }
      })
    }

    function getOldModulesPresets(targetModuleTypes, oldModulesPresetsValues) {
      var oldModuleTypeIndex = findModuleTypesRowIndex(targetModuleTypes, oldModulesPresetsValues)
      var oldModules = {}
      targetModuleTypes.forEach(function(moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1
        if (typeof rowIdx === "undefined") return
        oldModules[moduleType] = {}
        var row = oldModulesPresetsValues[rowIdx]
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Module Name") {
            var presetName = String(oldModulesPresetsValues[rowIdx - 1][col]).trim()
            var moduleName = String(oldModulesPresetsValues[rowIdx + 1][col]).trim()
            if (presetName && moduleName) {
              oldModules[moduleType][presetName] = moduleName
            }
          }
        }
      })
      return oldModules
    }

    function updateModulesInventory(targetModuleTypes, newModulesInventorySheet, oldModulesInventory) {
      var newModuleInventoryValues = newModulesInventorySheet.getDataRange().getValues()
      var newModuleTypeIndex = findModuleTypesRowIndex(targetModuleTypes, newModuleInventoryValues)
      targetModuleTypes.forEach(function(moduleType) {
        if (oldModulesInventory.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType]
          if (typeof rowIdx === "undefined") return
          var row = newModuleInventoryValues[rowIdx]
          var maxLevel = oldModulesInventory[moduleType]["Highest Level"] || 0
          var highestLevelCol = newModuleInventoryValues[rowIdx + 1].indexOf("Highest Level")
          for (var col = 1; col < row.length; col++) {
            var cellValue = String(row[col])
            if (cellValue.trim() !== "" && oldModulesInventory[moduleType].hasOwnProperty(cellValue)) {
              maxLevel = Math.max(maxLevel, oldModulesInventory[moduleType][cellValue].level)
              newModulesInventorySheet.getRange(rowIdx + 3, col + 1).setValue(oldModulesInventory[moduleType][cellValue].rarity)
              if (!newModulesInventorySheet.getRange(rowIdx + 3, col + 2).getFormula()) {
                newModulesInventorySheet.getRange(rowIdx + 3, col + 2).setValue(oldModulesInventory[moduleType][cellValue].level)
              }
              var substats = oldModulesInventory[moduleType][cellValue].substats
              var numRows = substats.length
              if (numRows > 0) {
                var numCols = substats[0].length
                newModulesInventorySheet.getRange(rowIdx + 5, col + 1, numRows, numCols).setValues(substats)
              }
            }
          }
          
          if (highestLevelCol !== -1) {
            newModulesInventorySheet.getRange(rowIdx + 3, highestLevelCol + 1).setValue(maxLevel)
          }
        }
      })
    }

    function getOldModulesInventory(targetModuleTypes, oldModulesInventoryValues) {
      var oldModuleTypeIndex = findModuleTypesRowIndex(targetModuleTypes, oldModulesInventoryValues)
      var oldModules = {}
      targetModuleTypes.forEach(function(moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType]
        if (typeof rowIdx === "undefined") return
        oldModules[moduleType] = {}
        var row = oldModulesInventoryValues[rowIdx]
        var highestLevelCol = oldModulesInventoryValues[rowIdx + 1].indexOf("Highest Level")
        if (highestLevelCol !== -1) {
          oldModules[moduleType]["Highest Level"] = oldModulesInventoryValues[rowIdx + 2][highestLevelCol]
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = String(row[col])
          if (cellValue.trim() !== "") {
            var moduleName = cellValue
            if (moduleName) {
              var removedRarity = ["Common", "Rare", "Rare+"]
              var moduleRarity = String(oldModulesInventoryValues[rowIdx + 2][col]).trim()
              if (removedRarity.includes(moduleRarity)) {
                moduleRarity = "Epic"
              }
              var moduleLevel = String(oldModulesInventoryValues[rowIdx + 2][col + 1]).trim()
              oldModules[moduleType][moduleName] = {"rarity": moduleRarity, "level": moduleLevel, "substats": []}
              for (var substat = rowIdx + 4; oldModulesInventoryValues.length; substat++) {
                var substatRow = oldModulesInventoryValues[substat]
                if (substatRow[col].trim() === "" && substatRow[col + 1].trim() === "") {
                  break
                }
                var substats = [substatRow[col], substatRow[col + 1]]
                oldModules[moduleType][moduleName]["substats"].push(substats)
              }
            }
          }
        }
      })
      return oldModules
    }

    function findModuleTypesRowIndex(targetModuleTypes, moduleRange) {
      var moduleTypeIndex = {}
      var moduleFound = {}
      targetModuleTypes.forEach(function(moduleType) {
        moduleFound[moduleType] = false
      })

      for (var i = 0; i < moduleRange.length; i++) {
        var cellValue = String(moduleRange[i][0]).toLowerCase()
        targetModuleTypes.forEach(function(moduleType) {
          if (!moduleFound[moduleType] && cellValue && cellValue.indexOf(moduleType) !== -1) {
            moduleTypeIndex[moduleType] = i
            moduleFound[moduleType] = true
          }
        })
        // If all terms are found, we can break early
        if (Object.values(moduleFound).every(Boolean)) {
          break
        }
      }

      return moduleTypeIndex
    }
  },

  convertVersion3: function(oldModulesSpreadsheet) {
    var targetModuleTypes = ["cannon", "armor", "generator", "core"]
    var oldModulesPresetsSheet = oldModulesSpreadsheet.getSheetByName("Modules Presets")
    var oldModulesPresetsValues = oldModulesPresetsSheet.getDataRange().getValues()
    
    var oldModuleTypeIndex = findModuleTypesRowIndex(targetModuleTypes, oldModulesPresetsValues)
    var oldModulesPresets = {}
    var oldModulesInventory = {}
    targetModuleTypes.forEach(function(moduleType) {
      var rowIdx = oldModuleTypeIndex[moduleType] + 1
      if (typeof rowIdx === "undefined") return
      oldModulesPresets[moduleType] = {}
      oldModulesInventory[moduleType] = {}
      var row = oldModulesPresetsValues[rowIdx]
      for (var col = 0; col < row.length; col++) {
        if (String(row[col]).trim() === "Module Name") {
          var moduleName = String(oldModulesPresetsValues[rowIdx + 1][col]).trim()
          if (moduleName) {
          var presetName = String(oldModulesPresetsValues[rowIdx - 1][col]).trim()
            if (presetName) {
              oldModulesPresets[moduleType][presetName] = moduleName
            }
            if (!oldModulesInventory[moduleType].hasOwnProperty(moduleName)) {
              var moduleRarity = String(oldModulesPresetsValues[rowIdx + 1][col + 1]).trim()
              var moduleLevel = String(oldModulesPresetsValues[rowIdx + 1][col + 2]).trim()
              oldModulesInventory[moduleType][moduleName] = {"rarity": moduleRarity, "level": moduleLevel, "substats": []}
              for (var substat = rowIdx + 3; oldModulesPresetsValues.length; substat++) {
                var substatRow = oldModulesPresetsValues[substat]
                if (substatRow[col].trim() === "" && substatRow[col + 1].trim() === "") {
                  break
                }
                var substats = [substatRow[col], substatRow[col + 1]]
                oldModulesInventory[moduleType][moduleName]["substats"].push(substats)
              }
            }
          }
        }
      }
    })
    return {"oldModulesInventory": oldModulesInventory, "oldModulesPresets": oldModulesPresets}
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion]
  },
}