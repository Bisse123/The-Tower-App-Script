const themesAndRelics = {

  sheetType: "Themes, Songs & Relics",

  /**
   * Reads Themes_Songs_Relics data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("themesAndRelics.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Themes_Songs_Relics data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: themesAndRelics.importData");
      const themesSheetName = "Themes & Songs";
      const relicsSheetName = "Relics";

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

      shared.addIDUpdatesToBatch(
        batchUpdate,
        this.sheetType,
        newSheetID,
        idsData,
        data.idMasterID,
      );

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
      var errorReport = errors.report("themesAndRelics.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes Themes into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldThemesNames
   * @param {Object} newThemesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
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

      var headerLocations = {};
      var batchUpdate = [];

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

            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = [];
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 });
          }
        }
      }

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
      var errorReport = errors.report("themesAndRelics.updateThemes", error, {
        sheetName: sheetName,
        oldThemesNames: oldThemesNames,
        newThemesData: newThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes Relics into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldRelics
   * @param {Object} newRelicsData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
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
      var errorReport = errors.report("themesAndRelics.updateRelics", error, {
        sheetName: sheetName,
        oldRelics: oldRelics,
        newRelicsData: newRelicsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Themes_Songs_Relics data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: themesAndRelics.version4_0");

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
      var errorReport = errors.report("themesAndRelics.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Themes from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldThemesData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
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

      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];
          if (oldThemeUnlocked === "Auto-fill from Player and Stuff") {
            oldThemesNames["autoFill"] = oldThemesData[row + 1][col];
            continue;
          }

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
      var errorReport = errors.report("themesAndRelics.getVersion4_0Themes", error, {
        oldThemesData: oldThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Relics from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldRelicsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0Relics: function (oldRelicsData) {
    try {
      console.log("Called: relics.getVersion4_0Relics");
      var oldRelicHeaderRow = -1;
      var relicNameIndex = -1;
      var relicUnlockedIndex = -1;

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
      var errorReport = errors.report("themesAndRelics.getVersion4_0Relics", error, {
        oldRelicsData: oldRelicsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Parses Themes_Songs_Relics data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseThemesAndRelicsData: function (data) {
    try {
      var towerSkins = {
        83: "Automaton",
        82: "Fairy",
        81: "Baby Dino",
        77: "Seahorse",
        76: "Meteorite",
        75: "Cake",
        74: "Rabbit In Hat",
        73: "Brain",
        72: "Bear",
        71: "Space Telescope",
        70: "Glitch",
        69: "Heart",
        68: "Shining Star",
        67: "Restless Eye",
        66: "Balloon",
        65: "Crystal",
        64: "Flying Car",
        63: "Pixel Soldier",
        59: "4th Anniversary",
        58: "Dj",
        57: "Frog",
        56: "Cthulhu",
        55: "Marshmallow",
        54: "Mech Warrior",
        53: "Neon Pi",
        52: "Crown",
        51: "Pocket Watch",
        50: "Black Hole",
        49: "Black Cat",
        48: "Snowman",
        47: "Unlucky Cow",
        43: "Noise Tower",
        42: "Umbrella",
        41: "Storm Eye",
        40: "Fisherman",
        39: "Elite Tower",
        38: "Starship",
        37: "Dive Helmet",
        36: "Dark Tower",
        35: "Toast Glass",
        34: "Invader",
        33: "Autumn Leaf",
        32: "Pumpkin",
        31: "Hourglass",
        30: "Virus",
        29: "Howling Wolf",
        28: "Sentinel",
        27: "Spider",
        24: "Prisma",
        23: "Neo Turbo",
        22: "Bunny",
        8: "Cherry Blossom",
        7: "Water Droplet",
        6: "Alien",
        5: "North Spirit",
        4: "Bee",
        3: "Plasma Ball",
        2: "Eye of the Lord",
        1: "Star",
      };
      var milestoneSkins = {
        80: "Cosmic",
        79: "Stellar",
        78: "Vortex",
        62: "Eclipse",
        61: "Cyber",
        60: "Atomic",
        46: "Rhino",
        45: "Dragon",
        44: "Cactus",
        26: "Tech Tree",
        25: "Panda",
        21: "Skull",
        20: "Cat",
        19: "Creepy Clown",
        18: "Cheese",
        17: "Turtle",
        16: "Mush-mush",
        15: "Fried Egg",
        14: "Sheep",
        13: "Butterfly",
        12: "Smile",
        11: "Yin-Yang",
        10: "Donut",
        9: "Shuriken",
      };
      var backgroundSkins = {
        58: "Steampunk",
        57: "Fairy Forest",
        56: "Jurassic Forest",
        55: "Coral Reef",
        54: "Meteor Shower",
        53: "5th Anniversary",
        52: "Magician",
        51: "Neuron",
        50: "Claw Machine",
        49: "Supernova",
        48: "Glitch",
        47: "Valentine",
        46: "Cozy Cosmos",
        45: "Crimson Horror",
        44: "Amusement Park",
        43: "Crystal Cave",
        42: "Cyberpunk",
        41: "Pixel Alien War",
        40: "Party",
        39: "Koi Pond",
        38: "Cthulhu",
        37: "Camping",
        36: "Mech World",
        35: "Pi Disk",
        34: "Throne Room",
        33: "Clock Tower",
        32: "Event Horizon",
        31: "Forest of Cats",
        30: "Snowstorm",
        29: "Abduction",
        28: "TV Wall",
        27: "Rainfall",
        26: "Hurricane",
        25: "Sunset River",
        24: "Invasion",
        23: "Hyper Space",
        22: "Deep Sea",
        21: "Dark Strands",
        20: "New Year",
        19: "Arcade",
        18: "Autumn Forest",
        17: "Haunted House",
        16: "Sand Storm",
        15: "Virus Field",
        14: "Mountain Night",
        13: "Matrix",
        12: "Cobweb",
        11: "Prismatic Lines",
        10: "Retrowave",
        9: "Easter",
        8: "Sakura",
        7: "Ocean Night",
        6: "Alien Ship",
        5: "Aurora",
        4: "Honeycomb",
        3: "Plasma Field",
        2: "Volcano",
        1: "Interstellar",
      };
      var guardianSkins = {
        22: "Pal",
        21: "Dewdrop",
        20: "Waddles",
        19: "Hermie",
        18: "Disco",
        17: "Shelly",
        16: "Mei",
        15: "Earl",
        14: "Frank",
        13: "Arwing",
        12: "Gaia",
        11: "Mickey",
        10: "Silk",
        9: "Iris",
        8: "Zepe",
        7: "Glenn",
        6: "Rolo",
        5: "Nyra",
        4: "Finn",
        2: "Muse",
        1: "Butter",
      };
      var profileBanners = {
        12: "Fairy Forest",
        11: "Coral Reef",
        10: "Magician",
        9: "Claw Machine",
        8: "Supernova",
        7: "Cosy Cosmos",
        6: "Crimson Horror",
        5: "Pixel Alien War",
        4: "Party",
        3: "Mech World",
        2: "What Time Is It Banner",
        1: "Arcade Banner",
      };
      var menuThemes = {
        11: "Fairy Forest",
        10: "Coral Reef",
        9: "Magician",
        8: "Claw Machine",
        7: "Supernova",
        6: "Cosy Cosmos",
        5: "Crimson Horror",
        4: "Pixel Alien War",
        3: "Party",
        2: "Mech World",
        1: "Dark Being",
      };
      var songs = {
        8: "Krisu - Forest Bathing",
        7: "Krisu - Hiding in Himalaya",
        6: "Krisu - Oceans Sings",
      };
      var relics = {
        312: "Fantasy Zeppelin",
        311: "Gadget Glasses",
        310: "Gear Heart",
        309: "Automated Pen",
        306: "Elder Staff",
        305: "Magic Elixir",
        304: "Pearl Shell",
        303: "Manta Ray",
        302: "T: XXIV Cosmic",
        301: "T: XXIII Stellar",
        300: "T: XXII Vortex",
        299: "Celestial Fishbones",
        298: "Cat Tome",
        297: "Galactic Beverage",
        296: "Starlight Yarn",
        295: "Grand Pyramid",
        294: "Ancient Writing",
        293: "Ancient Knowledge",
        292: "Ancient Art",
        291: "Tyrant's Skull",
        290: "Spiral Nautilus",
        289: "Hunter's Realm",
        288: "Ancient Footprint",
        287: "Rogue Planet",
        286: "Binary System",
        285: "Space Nebula",
        284: "Flying Object",
        283: "Light Scattering",
        282: "Light Spectrum",
        281: "Rainbow",
        280: "Prismatic Star",
        279: "Asteroid Belt",
        278: "Precious Minerals",
        277: "Meteor Impact",
        276: "Mining Drone",
        275: "Celebration",
        274: "Big Party",
        273: "Dangerous Tricks",
        272: "Magic Cards",
        271: "Geological Activity",
        270: "Obsidian",
        269: "New Island",
        268: "Magma River",
        267: "Global Threat",
        266: "Immunization",
        265: "Personal Care",
        264: "Viral Infection",
        263: "Body Control",
        262: "Neural Network",
        261: "Brain Net",
        260: "Synapse",
        259: "Storm Planet",
        258: "Big Tornado",
        257: "Natural Fire",
        256: "Nature's Fury",
        255: "Collector's Spirit",
        254: "Perfect Catch",
        253: "Quasar",
        252: "Elemental Explosion",
        251: "Instability",
        250: "Digital Disaster",
        249: "Research Object",
        248: "Broken Security",
        247: "Tori",
        246: "Forest Temple",
        245: "Ramen",
        244: "Festival Lanterns",
        243: "Moonlight",
        242: "Sailing At Night",
        241: "Night Shark",
        240: "Lighthouse",
        239: "Time Travel",
        238: "Clock Tower",
        237: "Space Distortion",
        236: "Ancient Times",
        235: "Star Planet",
        234: "Star Path",
        233: "Glimpse of Despair",
        232: "Blood Monster",
        231: "Alien Implants",
        230: "Crop Circles",
        229: "Alien Experiment",
        228: "Sudden Attack",
        227: "Cosmic Impact",
        226: "Northern Mountains",
        225: "Solar Flare",
        224: "Sky's Curtain",
        223: "Gift box",
        222: "Firework Rocket",
        221: "Champagne",
        220: "Party Popper",
        219: "Christmas Wreath",
        218: "Snowflake",
        217: "Winter Gloves",
        216: "Snow Globe",
        215: "Enemies",
        214: "Let's Play",
        213: "To Infinity",
        212: "Pinball",
        211: "Spider Forest",
        210: "Spider Poison",
        209: "Spider Vision",
        208: "Good Hunting",
        207: "Pierced Heart",
        206: "Lovely Gift",
        205: "Love Letter",
        204: "Bouquet",
        203: "Carousel Of Joy",
        202: "Amazing Prizes",
        201: "Delicious Food",
        200: "Happiness Balloons",
        199: "Full Minecart",
        198: "Crystals Bag",
        197: "Miner's Tool",
        196: "Explorer's Helmet",
        195: "Cybernetics",
        194: "Tech Weapon",
        193: "Holographic Ads",
        192: "Vr",
        191: "Brave Heroes",
        190: "World Domination",
        189: "T: XXI Eclipse",
        188: "T: XX Cyber",
        187: "T: XIX Atomic",
        186: "Night Life",
        185: "Let's Mix",
        184: "Warm Clothes",
        183: "Glowing Mushrooms",
        182: "Dry leaves",
        181: "Brunch",
        180: "Antenna",
        179: "No Signal",
        178: "Globalization",
        177: "Breaking News",
        176: "Fake Reality",
        175: "Tower Agent",
        174: "Gnosis",
        173: "Model Training",
        172: "River Of Plenty",
        171: "Good Catch",
        170: "Sunset Boat",
        169: "Fisherman Set",
        168: "Night City",
        167: "Retro Camera",
        166: "Magic Cube",
        165: "Floppy Disk",
        164: "Plasma Chamber",
        163: "Plasma Cell",
        162: "Plasma Vortex",
        161: "Plasma Globe",
        160: "Lilies",
        159: "Wind",
        158: "Grass",
        157: "Duck",
        156: "The Queen",
        155: "Honey Society",
        154: "Heavenly Sweet",
        153: "Honey Jar",
        152: "Omniscience",
        151: "Cosmic Freedom",
        150: "Madness Induced",
        149: "Rlyeh",
        148: "Nature's Wrath",
        147: "Eternal Quest",
        146: "Shining Light",
        145: "Safe Path",
        144: "Mech Head",
        143: "Fancy Wires",
        142: "Psychohistorian Brain",
        141: "Pi Seal",
        140: "Do While True",
        139: "Infinite Ruler",
        138: "Magic Egg",
        137: "Mystic Bunny",
        136: "Candy Core",
        135: "Bloom Burst",
        134: "Crown",
        133: "Throne",
        132: "Monolith",
        131: "Abduction Signal",
        130: "Photon Blade",
        129: "Quantum Drive",
        128: "Cursed Candle",
        127: "Whispering Web",
        126: "Time Compass",
        125: "Hourglass",
        124: "Alien Egg",
        123: "UFO Beam",
        122: "Light Speedometer",
        121: "Pulsar Core",
        120: "Dream Clock",
        119: "Temporal Rift",
        118: "Shadow Puppet",
        117: "Haunted Mirror",
        116: "Angler Fish",
        115: "Coral Crown",
        114: "3 Body Solution",
        113: "Falling Apple",
        112: "Party Mask",
        111: "Confetti Ball",
        110: "Pet Cat",
        109: "Lunar Cat Paw",
        108: "Arcade Token",
        107: "Power Glove",
        106: "Bonsai Tree",
        105: "Koi Fish",
        104: "Sleigh Bell",
        103: "Icicle",
        102: "Legend Badge",
        101: "Crop Circle",
        100: "Abduction Room",
        99: "Witch Hat",
        98: "Cauldron",
        97: "Scarf",
        96: "Acorn",
        95: "Kimono",
        94: "Tea Ceremony",
        93: "Neon Sunglasses",
        92: "Cassette",
        91: "Ash Cloud",
        90: "Lava Flow",
        89: "Planetary Rings",
        88: "Comet",
        87: "T: XVIII Singularity",
        86: "T: XVII Nebula",
        85: "T: XVI Quantum",
        84: "Cathode Ray Tube",
        83: "Remote Control",
        82: "6th Tower Birthday",
        81: "5th Tower Birthday",
        80: "4th Tower Birthday",
        79: "Sphinx",
        78: "Anubis",
        77: "Outbreak",
        76: "Rabies",
        75: "Cloud Lightning",
        74: "Rain Jacket",
        73: "Flying House",
        72: "Gale Winds",
        71: "Fish",
        70: "Hook",
        69: "Mountain Goat",
        68: "Summit Starlight",
        67: "Code Stream",
        66: "Clip Ons",
        65: "The Fly",
        64: "Cobweb",
        63: "Prismatic Shard",
        62: "Refraction Array",
        61: "Illuminati",
        60: "Pizza",
        59: "Wave",
        58: "Barnacle",
        57: "Star Ship",
        56: "Warp Gate",
        55: "The Kraken",
        54: "Submarine",
        53: "Creepy Smile",
        52: "Dark Sight",
        51: "Pixel Cube Heart",
        50: "Palm Tree",
        49: "Cheers",
        48: "Firework",
        47: "Controller",
        46: "Game Joystick",
        45: "Pumpkin",
        44: "Tower Latte",
        43: "Sakura Lantern",
        42: "Cherry",
        41: "Man Skull",
        40: "Spooky Bat",
        39: "Space Sundial",
        38: "Ancient Tome",
        37: "Alien Warp Drive",
        36: "Alien Head",
        35: "Contained Ions",
        34: "Aurora Vortex",
        33: "Stinger",
        32: "Honey Drop",
        31: "Plasma Arc",
        30: "Ionized Plasma",
        29: "Neuron",
        28: "Bacteriophage",
        27: "Spirit Wolf",
        26: "Dreamcatcher",
        25: "3rd Tower Birthday",
        24: "2nd Tower Birthday",
        23: "1st Tower Birthday",
        22: "T: XV Celestial",
        21: "T: XIV Arcane",
        20: "T: XIII Hyper",
        19: "T: XII Chrono",
        18: "T: XI Resonance",
        17: "T: X Plasma",
        16: "T: IX Fusion",
        15: "T: VIII Graviton",
        14: "T: VII Aether",
        13: "T: VI Nova",
        12: "T: V Ether",
        11: "T: IV Harmonic",
        10: "T: III Pulse",
        9: "T: II Lumin",
        8: "T: I Flux",
        7: "Tower Master",
        6: "Champion Badge",
        5: "Platinum Badge",
        4: "Gold Badge",
        3: "Silver Badge",
        2: "Copper Badge",
        1: "Red Pill",
        0: "No Spoon",
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
        if (!isUnlocked || !index) return;
        var towerName = towerSkins[index];
        var milestoneName = milestoneSkins[index];
        if (!towerName && !milestoneName) {
          console.log(`Warning: No tower or milestone skin name found for index ${index}`);
          towerName = `Unknown Tower/Milestone Skin ${index}`;
        }
        if (towerName) {
          oldThemesNames["Tower Skin"].push(towerName);
        }
        if (milestoneName) {
          oldThemesNames["Milestone Skin"].push(milestoneName);
        }
      });

      backgroundSkinsData.forEach(function (isUnlocked, index) {
        if (!isUnlocked || !index) return;
        var backgroundName = backgroundSkins[index];
        if (!backgroundName) {
          console.log(`Warning: No background skin name found for index ${index}`);
          backgroundName = `Unknown Background Skin ${index}`;
        }
        oldThemesNames["Background Skin"].push(backgroundName);
      });

      menuSkinsData.forEach(function (isUnlocked, index) {
        if (!isUnlocked || !index) return;
        var menuName = menuThemes[index];
        if (!menuName) {
          console.log(`Warning: No menu theme name found for index ${index}`);
          menuName = `Unknown Menu Theme ${index}`;
        }
        oldThemesNames["Menu"].push(menuName);
      });

      guardianSkinsData.forEach(function (isUnlocked, index) {
        if (!isUnlocked || !index) return;
        var guardianName = guardianSkins[index];
        if (!guardianName) {
          console.log(`Warning: No guardian skin name found for index ${index}`);
          guardianName = `Unknown Guardian Skin ${index}`;
        }
        oldThemesNames["Guardians"].push(guardianName);
      });

      profileBannersData.forEach(function (isUnlocked, index) {
        if (!isUnlocked || !index) return;
        var bannerName = profileBanners[index];
        if (!bannerName) {
          console.log(`Warning: No profile banner name found for index ${index}`);
          bannerName = `Unknown Profile Banner ${index}`;
        }
        oldThemesNames["Profile Banner"].push(bannerName);
      });

      songsData.forEach(function (isUnlocked, index) {
        if (!isUnlocked || index < 6) return;
        var songName = songs[index];
        if (!songName) {
          console.log(`Warning: No song name found for index ${index}`);
          songName = `Unknown Song ${index}`;
        }
        oldThemesNames["Songs"].push(songName);
      });

      relicsData.forEach(function (isUnlocked, index) {
        if (!isUnlocked) return;
        var relicName = relics[index];
        if (!relicName) {
          console.log(`Warning: No relic name found for index ${index}`);
          relicName = `Unknown Relic ${index}`;
        }
        oldRelics.push(relicName);
      });

      return {
        success: true,
        oldRelics: oldRelics,
        oldThemesNames: oldThemesNames,
        themesOrder: Object.keys(oldThemesNames),
      };
    } catch (error) {
      var errorReport = errors.report("themesAndRelics.parseThemesAndRelicsData", error, {
        data: data,
        oldThemesNames: oldThemesNames,
        oldRelics: oldRelics,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.0": this.version4_0.bind(this),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
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

};
