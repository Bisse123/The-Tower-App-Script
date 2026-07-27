const relics = {
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: relics.exportData");
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
        message: "Relics export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting relics data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: relics.importData");
      var newSpreadsheet = spreadsheets("Relics newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var requiredRanges = ["Relics", "IDS"];
      var newRelicsBatchResult = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges,
      );
      if (!newRelicsBatchResult || newRelicsBatchResult.length === 0) {
        console.log("Error getting relics sheet data");
        return {
          success: false,
          message: "Error getting relics sheet data",
        };
      }

      var newRelicsData = newRelicsBatchResult[0].values;
      var idsData = newRelicsBatchResult[1].values;

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

      // Only update relics if key exists
      if (data.hasOwnProperty("oldRelics")) {
        var oldRelics = data.oldRelics;
        var relicsResult = this.updateRelics(
          "Relics",
          oldRelics,
          newRelicsData,
        );
        if (!relicsResult || !relicsResult.success) {
          console.log(`Error updating relics: ${relicsResult.message}`);
          return relicsResult;
        }
        batchUpdate = batchUpdate.concat(relicsResult.batchUpdate || []);
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
        "Relics",
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
        message: `Relics import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing relics data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
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
  // #region Convert Versions
  version1_0: function () {
    try {
      console.log("Called: relics.version1_0");
      var oldSpreadsheet = spreadsheets("Relics oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Check if Relics sheet exists in old spreadsheet
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "Relics")) {
        console.log("Relics sheet not found in old relic spreadsheet");
        return {
          success: false,
          message: `Relics sheet not found in old relic spreadsheet™`,
        };
      }

      var oldRelicsBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        "Relics",
      ]);
      if (
        !oldRelicsBatchResult ||
        oldRelicsBatchResult.length === 0 ||
        !oldRelicsBatchResult[0].values
      ) {
        console.log(`Could not read data from old Relics sheet`);
        return {
          success: false,
          message: `Could not read data from old Relics sheet`,
        };
      }
      var oldRelicsData = oldRelicsBatchResult[0].values;

      var relicsData = this.getVersion1_0Relics(oldRelicsData);
      return relicsData;
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Relics
  getVersion1_0Relics: function (oldRelicsData) {
    try {
      console.log("Called: relics.getVersion1_0Relics");
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
      console.log("Error in getVersion1_0Relics: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Relics: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseRelicsData: function (data) {
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
      219: "Wreath",
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
      288: "Manta Ray",
      289: "Pearl Shell",
      290: "T:  XXII Vortex",
      291: "T:  XXIII Stellar",
      292: "T:  XXIV Cosmic",
    };

    const relicsData = data.relicsUnlocked || [];

    var oldRelics = [];
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
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
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
