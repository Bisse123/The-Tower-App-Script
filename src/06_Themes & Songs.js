const themes = {
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: themes.exportData");
      var getVersionFunction = this.convertVersionFunctions[versionDifference];
      if (!getVersionFunction) {
        console.log(`Unsupported version: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version: ${versionDifference}`,
        };
      }

      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(`${oldDataResult.message}`);
        return oldDataResult;
      }

      return {
        success: true,
        message: "Themes & Songs export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting themes data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: themes.importData");
      var newSpreadsheet = spreadsheets("Themes & Songs newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      // Batch get required data for update function only
      var requiredRanges = ["Themes & Songs", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var newThemesData = batchResults[0].values;
      var idsData = batchResults[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData,
      );
      if (
        !newSheetInfo ||
        !newSheetInfo.importStatus ||
        !newSheetInfo.importStatus.range
      ) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var batchUpdate = [];

      // Only update themes if key exists
      if (data.hasOwnProperty("oldThemesNames")) {
        var oldThemesNames = data.oldThemesNames;
        var themesResult = this.updateThemes(
          "Themes & Songs",
          oldThemesNames,
          newThemesData,
        );
        if (!themesResult || !themesResult.success) {
          console.log(`Error updating themes: ${themesResult.message}`);
          return themesResult;
        }
        batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Themes & Songs",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      // Apply all updates (including ID setting and import status)
      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        console.log(`Error applying batch updates to new spreadsheet`);
        return {
          success: false,
          message: "Error applying batch updates to new spreadsheet™",
        };
      }

      return {
        success: true,
        message: `Themes & Songs import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing themes data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateThemes: function (sheetName, oldThemesNames, newThemesData) {
    try {
      console.log("Called: themes.updateThemes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
      ];
      if (!newThemesData) {
        console.log(`Error getting new themes data`);
        return { success: false, message: "Error getting new themes data" };
      }
      var autoFill = false;
      if (oldThemesNames.hasOwnProperty("autoFill")) {
        autoFill =
          oldThemesNames.autoFill === true ||
          oldThemesNames.autoFill === "TRUE" ||
          oldThemesNames.autoFill === "true";
      }
      if (!autoFill) {
        targetThemes.push("Milestone Skin");
      }

      // For each header, store {col, startRow} for quick reference
      var headerLocations = {};
      var batchUpdate = [];

      // Pre-scan to find header columns and their start rows
      for (var i = 0; i < newThemesData.length; i++) {
        for (var j = 0; j < newThemesData[i].length; j++) {
          var newThemeUnlocked = String(newThemesData[i][j] || "").trim();
          if (newThemeUnlocked === "Auto-fill from Player and Stuff") {
            batchUpdate.push({
              range: `${sheetName}!${shared.columnToLetter(j + 1) + (i + 2)}`,
              values: [[autoFill]],
            });
            break;
          }
          if (targetThemes.indexOf(newThemeUnlocked) !== -1) {
            // If not already recorded for this col, store its location
            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = [];
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 }); // +1 to start below header
          }
        }
      }

      // For each header, possibly in multiple places
      for (var key in headerLocations) {
        headerLocations[key].forEach(function (loc) {
          var checkboxCol = loc.col;
          var startRow = loc.startRow;
          var checkedSet = new Set((oldThemesNames[key] || []).map(String));
          var checkboxArr = [];

          for (var row = startRow; row < newThemesData.length; row++) {
            var newThemeName = newThemesData[row][checkboxCol + 1];
            if (
              newThemeName === "" ||
              newThemeName === null ||
              typeof newThemeName === "undefined" ||
              targetThemes.indexOf(String(newThemeName || "").trim()) !== -1
            ) {
              break;
            }
            checkboxArr.push([checkedSet.has(String(newThemeName).trim())]);
          }

          if (checkboxArr.length > 0) {
            var startCell =
              shared.columnToLetter(checkboxCol + 1) + (startRow + 1);
            var endCell =
              shared.columnToLetter(checkboxCol + 1) +
              (startRow + checkboxArr.length);
            batchUpdate.push({
              range: `${sheetName}!${startCell}:${endCell}`,
              values: checkboxArr,
            });
          }
        });
      }

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: "Themes & Songs updated successfully",
          batchUpdate: batchUpdate,
        };
      }
      return { success: true, message: "No updates needed for Themes & Songs" };
    } catch (error) {
      console.log("Error in updateThemes: " + error.toString());
      return {
        success: false,
        message: "Error in updateThemes: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version2_1_6: function () {
    try {
      console.log("Called: themes.version2_1_6");
      var oldSpreadsheet = spreadsheets("Themes & Songs oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var themesValuesRange = "Themes & Songs";
      var themesOldBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        themesValuesRange,
      ]);
      if (
        !themesOldBatchResult ||
        themesOldBatchResult.length === 0 ||
        !themesOldBatchResult[0].values
      ) {
        console.log(`Error getting old themes data`);
        return { success: false, message: "Error getting old themes data" };
      }
      var oldThemesData = themesOldBatchResult[0].values;

      var themesData = this.getVersion2_1_6Themes(oldThemesData);
      return themesData;
    } catch (error) {
      console.log("Error in version2_1_6: " + error.toString());
      return {
        success: false,
        message: "Error in version2_1_6: " + error.message,
      };
    }
  },

  version1_0: function () {
    try {
      console.log("Called: themes.version1_0");
      var oldSpreadsheet = spreadsheets("Themes & Songs oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var themesValuesRange = "Themes & Songs";
      var themesOldBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        themesValuesRange,
      ]);
      if (
        !themesOldBatchResult ||
        themesOldBatchResult.length === 0 ||
        !themesOldBatchResult[0].values
      ) {
        console.log(`Error getting old themes data`);
        return { success: false, message: "Error getting old themes data" };
      }
      var oldThemesData = themesOldBatchResult[0].values;

      var themesData = this.getVersion1_0Themes(oldThemesData);
      return themesData;
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Themes
  getVersion2_1_6Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion2_1_6Themes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
        "Milestone Skin",
      ];

      var oldThemesNames = {};

      targetThemes.forEach(function (header) {
        oldThemesNames[header] = [];
      });
      var currentHeader = null;
      var headerCol = -1;
      // Loop through each column first, then rows
      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];
          if (oldThemeUnlocked === "Auto-fill from Player and Stuff") {
            oldThemesNames["autoFill"] = oldThemesData[row + 1][col];
            continue;
          }
          // If cell is a header
          if (
            targetThemes.indexOf(String(oldThemeUnlocked || "").trim()) !== -1
          ) {
            currentHeader = String(oldThemeUnlocked || "").trim();
            headerCol = col;
            continue;
          }
          var isThemeUnlocked =
            oldThemeUnlocked === true ||
            oldThemeUnlocked === "TRUE" ||
            oldThemeUnlocked === "true";
          if (currentHeader && col === headerCol && isThemeUnlocked) {
            var oldThemeName = oldThemesData[row][col + 1];
            oldThemesNames[currentHeader].push(oldThemeName);
          }
        }
      }

      return {
        success: true,
        oldThemesNames: oldThemesNames,
      };
    } catch (error) {
      console.log("Error in getVersion2_1_6Themes: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_1_6Themes: " + error.message,
      };
    }
  },

  getVersion1_0Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion1_0Themes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
      ];

      var oldThemesNames = {};

      targetThemes.forEach(function (header) {
        oldThemesNames[header] = [];
      });
      var currentHeader = null;
      var headerCol = -1;
      // Loop through each column first, then rows
      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];
          // If cell is a header
          if (
            targetThemes.indexOf(String(oldThemeUnlocked || "").trim()) !== -1
          ) {
            currentHeader = String(oldThemeUnlocked || "").trim();
            if (oldThemesData[row][col + 2] === "Tier Unlocked") {
              currentHeader = "Milestone Skin";
              oldThemesNames["Milestone Skin"] = [];
            }
            headerCol = col;
            continue;
          }
          var isThemeUnlocked =
            oldThemeUnlocked === true ||
            oldThemeUnlocked === "TRUE" ||
            oldThemeUnlocked === "true";
          if (currentHeader && col === headerCol && isThemeUnlocked) {
            var oldThemeName = oldThemesData[row][col + 1];
            oldThemesNames[currentHeader].push(oldThemeName);
          }
        }
      }

      return {
        success: true,
        oldThemesNames: oldThemesNames,
      };
    } catch (error) {
      console.log("Error in getVersion1_0Themes: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Themes: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseThemesData: function (data) {
    var towerSkins = {
      1: "Star",
      2: "Eye of the Lord",
      3: "Plasma Ball",
      4: "Bee",
      5: "North Spirit",
      6: "Alien",
      7: "Water Droplet",
      8: "Cherry Blossom",
      22: "Bunny",
      23: "Neo Turbo",
      24: "Prisma",
      27: "Spider",
      28: "Sentinel",
      29: "Howling Wolf",
      30: "Virus",
      31: "Hourglass",
      32: "Pumpkin",
      33: "Autumn Leaf",
      34: "Invader",
      35: "Toast Glass",
      36: "Dark Tower",
      37: "Dive Helmet",
      38: "Starship",
      39: "Elite Tower",
      40: "Fisherman",
      41: "Storm Eye",
      42: "Umbrella",
      43: "Noise Tower",
      47: "Unlucky Cow",
      48: "Snowman",
      49: "Black Cat",
      50: "Black Hole",
      51: "Pocket Watch",
      52: "Crown",
      53: "Neon Pi",
      54: "Mech Warrior",
      55: "Marshmallow",
      56: "Cthulhu",
      57: "Frog",
      58: "Dj",
      59: "4th Anniversary",
      63: "Pixel Soldier",
      64: "Flying Car",
      65: "Crystal",
      66: "Balloon",
      67: "Restless Eye",
      68: "Shining Star",
      69: "Heart",
      70: "Glitch",
      71: "Space Telescope",
      72: "Bear",
      73: "Brain",
      74: "Rabbit In Hat",
      75: "Cake",
      76: "Meteorite",
      77: "Seahorse",
      81: "Baby Dino",
    };
    var milestoneSkins = {
      9: "Shuriken",
      10: "Donut",
      11: "Yin-Yang",
      12: "Smile",
      13: "Butterfly",
      14: "Sheep",
      15: "Fried Egg",
      16: "Mush-mush",
      17: "Turtle",
      18: "Cheese",
      19: "Creepy Clown",
      20: "Cat",
      21: "Skull",
      25: "Panda",
      26: "Tech Tree",
      44: "Cactus",
      45: "Dragon",
      46: "Rhino",
      60: "Atomic",
      61: "Cyber",
      62: "Eclipse",
      78: "Vortex",
      79: "Stellar",
      80: "Cosmic",
    };
    var backgroundSkins = {
      1: "Interstellar",
      2: "Volcano",
      3: "Plasma Field",
      4: "Honeycomb",
      5: "Aurora",
      6: "Alien Ship",
      7: "Ocean Night",
      8: "Sakura",
      9: "Easter",
      10: "Retrowave",
      11: "Prismatic Lines",
      12: "Cobweb",
      13: "Matrix",
      14: "Mountain Night",
      15: "Virus Field",
      16: "Sand Storm",
      17: "Haunted House",
      18: "Autumn Forest",
      19: "Arcade",
      20: "New Year",
      21: "Dark Strands",
      22: "Deep Sea",
      23: "Hyper Space",
      24: "Invasion",
      25: "Sunset River",
      26: "Hurricane",
      27: "Rainfall",
      28: "TV Wall",
      29: "Abduction",
      30: "Snowstorm",
      31: "Forest of Cats",
      32: "Event Horizon",
      33: "Clock Tower",
      34: "Throne Room",
      35: "Pi Disk",
      36: "Mech World",
      37: "Camping",
      38: "Cthulhu",
      39: "Koi Pond",
      40: "Party",
      41: "Pixel Alien War",
      42: "Cyberpunk",
      43: "Crystal Cave",
      44: "Amusement Park",
      45: "Crimson Horror",
      46: "Cozy Cosmos",
      47: "Valentine",
      48: "Glitch",
      49: "Supernova",
      50: "Claw Machine",
      51: "Neuron",
      52: "Magician",
      53: "5th Anniversary",
      54: "Meteor Shower",
      55: "Coral Reef",
      56: "Jurassic Forrest",
    };
    var guardianSkins = {
      1: "Butter",
      2: "Muse",
      4: "Finn",
      5: "Nyra",
      6: "Rolo",
      7: "Glenn",
      8: "Zepe",
      9: "Iris",
      10: "Silk",
      11: "Mickey",
      12: "Gaia",
      13: "Arwing",
      14: "Frank",
      15: "Earl",
      16: "Mei",
      17: "Shelly",
      18: "Disco",
      19: "Hermie",
      20: "Waddles",
    };
    var profileBanners = {
      1: "Arcade Banner",
      2: "What Time Is It Banner",
      3: "Mech World",
      4: "Party",
      5: "Pixel Alien War",
      6: "Crimson Horror",
      7: "Cosy Cosmos",
      8: "Supernova",
      9: "Claw Machine",
      10: "Magician",
      11: "Coral Reef",
    };
    var menuThemes = {
      1: "Dark Being",
      2: "Mech World",
      3: "Party",
      4: "Pixel Alien War",
      5: "Crimson Horror",
      6: "Cosy Cosmos",
      7: "Supernova",
      8: "Claw Machine",
      9: "Magician",
      10: "Coral Reef",
    };
    var songs = {
      6: "Krisu - Oceans Sings",
      7: "Krisu - Hiding in Himalaya",
      8: "Krisu - Forest Bathing",
    };

    const towerSkinsData = data.towerSkins || [];
    const backgroundSkinsData = data.backgroundSkins || [];
    const menuSkinsData = data.menuSkins || [];
    const guardianSkinsData = data.guardianSkins || [];
    const profileBannersData = data.profileBanners || [];
    const songsData = data.songs || [];

    var oldThemesNames = {
      "Tower Skin": [],
      "Background Skin": [],
      "Milestone Skin": [],
      Guardians: [],
      "Profile Banner": [],
      Menu: [],
      Songs: [],
    };

    towerSkinsData.forEach(function (isUnlocked, index) {
      if (!isUnlocked) return;
      if (towerSkins[index]) {
        oldThemesNames["Tower Skin"].push(towerSkins[index]);
      }
      if (milestoneSkins[index]) {
        oldThemesNames["Milestone Skin"].push(milestoneSkins[index]);
      }
    });

    backgroundSkinsData.forEach(function (isUnlocked, index) {
      if (isUnlocked && backgroundSkins[index]) {
        oldThemesNames["Background Skin"].push(backgroundSkins[index]);
      }
    });

    menuSkinsData.forEach(function (isUnlocked, index) {
      if (isUnlocked && menuThemes[index]) {
        oldThemesNames["Menu"].push(menuThemes[index]);
      }
    });

    guardianSkinsData.forEach(function (isUnlocked, index) {
      if (isUnlocked && guardianSkins[index]) {
        oldThemesNames["Guardians"].push(guardianSkins[index]);
      }
    });

    profileBannersData.forEach(function (isUnlocked, index) {
      if (isUnlocked && profileBanners[index]) {
        oldThemesNames["Profile Banner"].push(profileBanners[index]);
      }
    });

    songsData.forEach(function (isUnlocked, index) {
      if (isUnlocked && songs[index]) {
        oldThemesNames["Songs"].push(songs[index]);
      }
    });

    return {
      oldThemesNames: oldThemesNames,
      themesOrder: Object.keys(oldThemesNames),
    };
  },
  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.1.6": this.version2_1_6.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = Object.keys(this.convertVersionFunctions);

    var sortedThresholds = versionCompatibility.slice().sort(function (a, b) {
      return shared.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = shared.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },

  // #endregion
};
