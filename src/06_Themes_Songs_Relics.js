const themesAndRelics = {
  // #region Sheet References
  sheetType: "Themes, Songs & Relics",
  oldSpreadsheetName: "Themes, Songs & Relics oldSpreadsheet",
  newSpreadsheetName: "Themes, Songs & Relics newSpreadsheet",

  // #endregion
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: themesAndRelics.exportData");
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
        message: "Themes, Songs & Relics export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message:
          "Error exporting Themes, Songs & Relics data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: themesAndRelics.importData");
      const themesSheetName = "Themes & Songs";
      const relicsSheetName = "Relics";
      var newSpreadsheet = spreadsheets(this.newSpreadsheetName);
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      // Batch get required data for update function only
      var requiredRanges = [themesSheetName, relicsSheetName, "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length < requiredRanges.length) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var newThemesData = batchResults[0].values;
      var newRelicsData = batchResults[1].values;
      var idsData = batchResults[2].values;

      var batchUpdate = [];

      // Only update themes if key exists
      if (data.hasOwnProperty("oldThemesNames")) {
        var oldThemesNames = data.oldThemesNames;
        var themesResult = this.updateThemes(
          themesSheetName,
          oldThemesNames,
          newThemesData,
        );
        if (!themesResult || !themesResult.success) {
          console.log(`Error updating themes: ${themesResult.message}`);
          return themesResult;
        }
        batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
      }

      // Only update relics if key exists
      if (data.hasOwnProperty("oldRelics")) {
        var oldRelics = data.oldRelics;
        var relicsResult = this.updateRelics(
          relicsSheetName,
          oldRelics,
          newRelicsData,
        );
        if (!relicsResult || !relicsResult.success) {
          console.log(`Error updating relics: ${relicsResult.message}`);
          return relicsResult;
        }
        batchUpdate = batchUpdate.concat(relicsResult.batchUpdate || []);
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(
        batchUpdate,
        this.sheetType,
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
        message: `Themes, Songs & Relics import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing Themes, Songs & Relics data: ${error.message}`,
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

  updateRelics: function (sheetName, oldRelics, newRelicsData) {
    try {
      console.log("Called: relics.updateRelics");
      if (!newRelicsData || newRelicsData.length < 3) {
        console.log(`Not enough data in new Relics sheet`);
        return {
          success: false,
          message: `Not enough data in new Relics sheet`,
        };
      }

      var newRelicHeaderRow = null;
      var newRelicNameCol = null;
      var newRelicUnlockedCol = null;

      // Scan each row to find the header
      for (var row = 0; row < newRelicsData.length; row++) {
        var rowValues = newRelicsData[row];
        var relicNameIndex = rowValues.indexOf("Relic Name");
        var relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          newRelicHeaderRow = row + 1;
          newRelicNameCol = relicNameIndex + 1;
          newRelicUnlockedCol = relicUnlockedIndex + 1;
          break;
        }
      }

      if (!newRelicHeaderRow) {
        console.log(`Could not find header row in new Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in new Relics sheet`,
        };
      }

      var startRow = newRelicHeaderRow + 1;

      // Build unlocked status array directly by iterating through new relics data
      var newRelicsUnlocked = [];
      newRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = (row[newRelicNameCol - 1] || "").trim();
        if (String(relicName).trim() !== "") {
          if (oldRelics.includes(relicName)) {
            newRelicsUnlocked.push([true]);
          } else {
            newRelicsUnlocked.push([false]);
          }
        }
      });
      if (newRelicsUnlocked.length > 0) {
        var endRow = startRow + newRelicsUnlocked.length - 1;
        var unlockedRange = `${sheetName}!${shared.columnToLetter(
          newRelicUnlockedCol,
        )}${startRow}:${shared.columnToLetter(newRelicUnlockedCol)}${endRow}`;

        var batchUpdate = [
          {
            range: unlockedRange,
            values: newRelicsUnlocked,
          },
        ];
        return {
          success: true,
          message: `Relics updated successfully: ${newRelicsUnlocked.length} relics processed`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for relics`,
      };
    } catch (error) {
      console.log("Error in updateRelics: " + error.toString());
      return {
        success: false,
        message: "Error updating relics: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Version
  version4_0: function () {
    try {
      console.log("Called: themesAndRelics.version4_0");
      var oldSpreadsheet = spreadsheets(this.oldSpreadsheetName);
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet not found",
        };
      }
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var requiredRanges = ["Themes & Songs", "Relics"];
      var batchResults = SheetsAPI.batchGetValues(oldSheetID, requiredRanges);
      if (!batchResults || batchResults.length < requiredRanges.length) {
        console.log(`Could not read required data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from old spreadsheet",
        };
      }

      var oldThemesData = batchResults[0].values;
      var oldRelicsData = batchResults[1].values;
      
      var themesResult = this.getVersion4_0Themes(oldThemesData);
      if (!themesResult || !themesResult.success) {
        console.log(`Error converting themes: ${themesResult.message}`);
        return themesResult;
      }

      var relicsResult = this.getVersion4_0Relics(oldRelicsData);
      if (!relicsResult || !relicsResult.success) {
        console.log(`Error converting relics: ${relicsResult.message}`);
        return relicsResult;
      }

      return {
        success: true,
        oldThemesNames: themesResult.oldThemesNames,
        oldRelics: relicsResult.oldRelics,
      };
      

    } catch (error) {
      console.log(`Error in version4_0: ${error.toString()}`);
      return {
        success: false,
        message: `Error converting Themes, Songs & Relics data for version 4.0: ${error.message}`,
      };
    }
  },

  // #endregion
  // region getThemes
  getVersion4_0Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion4_0Themes");
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
      console.log("Error in getVersion4_0Themes: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0Themes: " + error.message,
      };
    }
  },

  // #endregion
  // region getRelics
  getVersion4_0Relics: function (oldRelicsData) {
    try {
      console.log("Called: relics.getVersion4_0Relics");
      var oldRelicHeaderRow = -1;
      var relicNameIndex = -1;
      var relicUnlockedIndex = -1;

      // Scan each row to find the header
      for (var row = 0; row < oldRelicsData.length; row++) {
        var rowValues = oldRelicsData[row];
        relicNameIndex = rowValues.indexOf("Relic Name");
        relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          oldRelicHeaderRow = row + 1;
          break;
        }
      }

      if (oldRelicHeaderRow === -1) {
        console.log(`Could not find header row in old Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in old Relics sheet`,
        };
      }

      var startRow = oldRelicHeaderRow + 1;

      var oldRelics = [];
      oldRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = row[relicNameIndex].trim();
        if (relicName.includes("T:")) {
          relicName = relicName.replace(/T:\s*/g, "T: ");
        }
        var isUnlocked = row[relicUnlockedIndex];

        if (
          relicName &&
          (isUnlocked === true ||
            isUnlocked === "TRUE" ||
            isUnlocked === "true")
        ) {
          oldRelics.push(relicName);
        }
      });

      return {
        success: true,
        oldRelics: oldRelics,
      };
    } catch (error) {
      console.log("Error in getVersion4_0Relics: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0Relics: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseThemesAndRelicsData: function (data) {
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
      56: "Jurassic Forest",
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
    var relics = {
      0: "No Spoon",
      1: "Red Pill",
      2: "Copper Badge",
      3: "Silver Badge",
      4: "Gold Badge",
      5: "Platinum Badge",
      6: "Champion Badge",
      7: "Tower Master",
      8: "T: I Flux",
      9: "T: II Lumin",
      10: "T: III Pulse",
      11: "T: IV Harmonic",
      12: "T: V Ether",
      13: "T: VI Nova",
      14: "T: VII Aether",
      15: "T: VIII Graviton",
      16: "T: IX Fusion",
      17: "T: X Plasma",
      18: "T: XI Resonance",
      19: "T: XII Chrono",
      20: "T: XIII Hyper",
      21: "T: XIV Arcane",
      22: "T: XV Celestial",
      23: "1st Tower Birthday",
      24: "2nd Tower Birthday",
      25: "3rd Tower Birthday",
      26: "Dreamcatcher",
      27: "Spirit Wolf",
      28: "Bacteriophage",
      29: "Neuron",
      30: "Ionized Plasma",
      31: "Plasma Arc",
      32: "Honey Drop",
      33: "Stinger",
      34: "Aurora Vortex",
      35: "Contained Ions",
      36: "Alien Head",
      37: "Alien Warp Drive",
      38: "Ancient Tome",
      39: "Space Sundial",
      40: "Spooky Bat",
      41: "Man Skull",
      42: "Cherry",
      43: "Sakura Lantern",
      44: "Tower Latte",
      45: "Pumpkin",
      46: "Game Joystick",
      47: "Controller",
      48: "Firework",
      49: "Cheers",
      50: "Palm Tree",
      51: "Pixel Cube Heart",
      52: "Dark Sight",
      53: "Creepy Smile",
      54: "Submarine",
      55: "The Kraken",
      56: "Warp Gate",
      57: "Star Ship",
      58: "Barnacle",
      59: "Wave",
      60: "Pizza",
      61: "Illuminati",
      62: "Refraction Array",
      63: "Prismatic Shard",
      64: "Cobweb",
      65: "The Fly",
      66: "Clip Ons",
      67: "Code Stream",
      68: "Summit Starlight",
      69: "Mountain Goat",
      70: "Hook",
      71: "Fish",
      72: "Gale Winds",
      73: "Flying House",
      74: "Rain Jacket",
      75: "Cloud Lightning",
      76: "Rabies",
      77: "Outbreak",
      78: "Anubis",
      79: "Sphinx",
      80: "4th Tower Birthday",
      81: "5th Tower Birthday",
      82: "6th Tower Birthday",
      83: "Remote Control",
      84: "Cathode Ray Tube",
      85: "T: XVI Quantum",
      86: "T: XVII Nebula",
      87: "T: XVIII Singularity",
      88: "Comet",
      89: "Planetary Rings",
      90: "Lava Flow",
      91: "Ash Cloud",
      92: "Cassette",
      93: "Neon Sunglasses",
      94: "Tea Ceremony",
      95: "Kimono",
      96: "Acorn",
      97: "Scarf",
      98: "Cauldron",
      99: "Witch Hat",
      100: "Abduction Room",
      101: "Crop Circle",
      102: "Legend Badge",
      103: "Icicle",
      104: "Sleigh Bell",
      105: "Koi Fish",
      106: "Bonsai Tree",
      107: "Power Glove",
      108: "Arcade Token",
      109: "Lunar Cat Paw",
      110: "Pet Cat",
      111: "Confetti Ball",
      112: "Party Mask",
      113: "Falling Apple",
      114: "3 Body Solution",
      115: "Coral Crown",
      116: "Angler Fish",
      117: "Haunted Mirror",
      118: "Shadow Puppet",
      119: "Temporal Rift",
      120: "Dream Clock",
      121: "Pulsar Core",
      122: "Light Speedometer",
      123: "UFO Beam",
      124: "Alien Egg",
      125: "Hourglass",
      126: "Time Compass",
      127: "Whispering Web",
      128: "Cursed Candle",
      129: "Quantum Drive",
      130: "Photon Blade",
      131: "Abduction Signal",
      132: "Monolith",
      133: "Throne",
      134: "Crown",
      135: "Bloom Burst",
      136: "Candy Core",
      137: "Mystic Bunny",
      138: "Magic Egg",
      139: "Infinite Ruler",
      140: "Do While True",
      141: "Pi Seal",
      142: "Psychohistorian Brain",
      143: "Fancy Wires",
      144: "Mech Head",
      145: "Safe Path",
      146: "Shining Light",
      147: "Eternal Quest",
      148: "Nature's Wrath",
      149: "Rlyeh",
      150: "Madness Induced",
      151: "Cosmic Freedom",
      152: "Omniscience",
      153: "Honey Jar",
      154: "Heavenly Sweet",
      155: "Honey Society",
      156: "The Queen",
      157: "Duck",
      158: "Grass",
      159: "Wind",
      160: "Lilies",
      161: "Plasma Globe",
      162: "Plasma Vortex",
      163: "Plasma Cell",
      164: "Plasma Chamber",
      165: "Floppy Disk",
      166: "Magic Cube",
      167: "Retro Camera",
      168: "Night City",
      169: "Fisherman Set",
      170: "Sunset Boat",
      171: "Good Catch",
      172: "River Of Plenty",
      173: "Model Training",
      174: "Gnosis",
      175: "Tower Agent",
      176: "Fake Reality",
      177: "Breaking News",
      178: "Globalization",
      179: "No Signal",
      180: "Antenna",
      181: "Brunch",
      182: "Dry leaves",
      183: "Glowing Mushrooms",
      184: "Warm Clothes",
      185: "Let's Mix",
      186: "Night Life",
      187: "T: XIX Atomic",
      188: "T: XX Cyber",
      189: "T: XXI Eclipse",
      190: "World Domination",
      191: "Brave Heroes",
      192: "Vr",
      193: "Holographic Ads",
      194: "Tech Weapon",
      195: "Cybernetics",
      196: "Explorer's Helmet",
      197: "Miner's Tool",
      198: "Crystals Bag",
      199: "Full Minecart",
      200: "Happiness Balloons",
      201: "Delicious Food",
      202: "Amazing Prizes",
      203: "Carousel Of Joy",
      204: "Bouquet",
      205: "Love Letter",
      206: "Lovely Gift",
      207: "Pierced Heart",
      208: "Good Hunting",
      209: "Spider Vision",
      210: "Spider Poison",
      211: "Spider Forest",
      212: "Pinball",
      213: "To Infinity",
      214: "Let's Play",
      215: "Enemies",
      216: "Snow Globe",
      217: "Winter Gloves",
      218: "Snowflake",
      219: "Christmas Wreath",
      220: "Party Popper",
      221: "Champagne",
      222: "Firework Rocket",
      223: "Gift box",
      224: "Sky's Curtain",
      225: "Solar Flare",
      226: "Northern Mountains",
      227: "Cosmic Impact",
      228: "Sudden Attack",
      229: "Alien Experiment",
      230: "Crop Circles",
      231: "Alien Implants",
      232: "Blood Monster",
      233: "Glimpse of Despair",
      234: "Star Path",
      235: "Star Planet",
      236: "Ancient Times",
      237: "Space Distortion",
      238: "Clock Tower",
      239: "Time Travel",
      240: "Lighthouse",
      241: "Night Shark",
      242: "Sailing At Night",
      243: "Moonlight",
      244: "Festival Lanterns",
      245: "Ramen",
      246: "Forest Temple",
      247: "Tori",
      248: "Broken Security",
      249: "Research Object",
      250: "Digital Disaster",
      251: "Instability",
      252: "Elemental Explosion",
      253: "Quasar",
      254: "Perfect Catch",
      255: "Collector's Spirit",
      256: "Nature's Fury",
      257: "Natural Fire",
      258: "Big Tornado",
      259: "Storm Planet",
      260: "Synapse",
      261: "Brain Net",
      262: "Neural Network",
      263: "Body Control",
      264: "Viral Infection",
      265: "Personal Care",
      266: "Immunization",
      267: "Global Threat",
      268: "Magma River",
      269: "New Island",
      270: "Obsidian",
      271: "Geological Activity",
      272: "Magic Cards",
      273: "Dangerous Tricks",
      274: "Big Party",
      275: "Celebration",
      276: "Mining Drone",
      277: "Meteor Impact",
      278: "Precious Minerals",
      279: "Asteroid Belt",
      280: "Prismatic Star",
      281: "Rainbow",
      282: "Light Spectrum",
      283: "Light Scattering",
      284: "Flying Object",
      285: "Space Nebula",
      286: "Binary System",
      287: "Rogue Planet",
      288: "Ancient Footprint",
      289: "Hunter's Realm",
      290: "Spiral Nautilus",
      291: "Tyrant's Skull",
      292: "Manta Ray",
      293: "Pearl Shell",
      294: "T: XXII Vortex",
      295: "T: XXIII Stellar",
      296: "T: XXIV Cosmic",
    };

    const towerSkinsData = data.towerSkins || [];
    const backgroundSkinsData = data.backgroundSkins || [];
    const menuSkinsData = data.menuSkins || [];
    const guardianSkinsData = data.guardianSkins || [];
    const profileBannersData = data.profileBanners || [];
    const songsData = data.songs || [];
    const relicsData = data.relicsUnlocked || [];

    var oldThemesNames = {
      "Tower Skin": [],
      "Background Skin": [],
      "Milestone Skin": [],
      Guardians: [],
      "Profile Banner": [],
      Menu: [],
      Songs: [],
    };
    
    var oldRelics = [];

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

    relicsData.forEach(function (relicStatus, index) {
      if (relicStatus < 1) {
        return;
      }
      var relicName = relics[index];
      if (!relicName) {
        return;
      }
      oldRelics.push(relicName);
    });

    return {
      oldRelics: oldRelics,
      oldThemesNames: oldThemesNames,
      themesOrder: Object.keys(oldThemesNames),
    };
  },
  
  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v4.0": this.version4_0.bind(this),
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
