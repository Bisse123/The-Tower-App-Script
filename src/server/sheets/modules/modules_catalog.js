const modulesCatalog = {

  /**
   * Module names and categories by their info index in the save file.
   */
  moduleNames: {
    7: { name: "Havoc Bringer", category: "Cannon" },
    8: { name: "Death Penalty", category: "Cannon" },
    9: { name: "Being Annihilator", category: "Cannon" },
    10: { name: "Astral Deliverance", category: "Cannon" },
    17: { name: "Wormhole Redirector", category: "Armor" },
    18: { name: "Negative Mass Projector", category: "Armor" },
    19: { name: "Space Displacer", category: "Armor" },
    20: { name: "Anti-Cube Portal", category: "Armor" },
    27: { name: "Black Hole Digestor", category: "Generator" },
    28: { name: "Pulsar Harvester", category: "Generator" },
    29: { name: "Galaxy Compressor", category: "Generator" },
    30: { name: "Singularity Harness", category: "Generator" },
    37: { name: "Multiverse Nexus", category: "Core" },
    38: { name: "Dimension Core", category: "Core" },
    39: { name: "Harmony Conductor", category: "Core" },
    40: { name: "Om Chip", category: "Core" },
    41: { name: "Shrink Ray", category: "Cannon" },
    42: { name: "Sharp Fortitude", category: "Armor" },
    43: { name: "Project Funding", category: "Generator" },
    44: { name: "Magnetic Hook", category: "Core" },
    45: { name: "Amplifying Strike", category: "Cannon" },
    46: { name: "Orbital Augment", category: "Armor" },
    47: { name: "Restorative Bonus", category: "Generator" },
    48: { name: "Primordial Collapse", category: "Core" },
    49: { name: "New Generator", category: "Generator" },
    50: { name: "Sentry Protocol", category: "Armor" },
    51: { name: "Gilded Sniper", category: "Cannon" },
    52: { name: "Tactical Barrage", category: "Core" },
  },

  /**
   * Module rarity names by their rarity number in the save file.
   */
  moduleRarities: {
    1: "Common",
    2: "Rare",
    3: "Rare+",
    4: "Epic",
    5: "Epic+",
    6: "Legendary",
    7: "Legendary+",
    8: "Mythic",
    9: "Mythic+",
    10: "Ancestral",
    11: "Ancestral 1*",
    12: "Ancestral 2*",
    13: "Ancestral 3*",
    14: "Ancestral 4*",
    15: "Ancestral 5*",
  },

  /**
   * Module category names by their index in the save file.
   */
  moduleCategories: {
    0: "Cannon",
    1: "Armor",
    2: "Generator",
    3: "Core",
  },

  /**
   * Substat rarity names, lowest first.
   */
  effectRarities: [
    "Common",
    "Rare",
    "Epic",
    "Legendary",
    "Mythic",
    "Ancestral",
  ],

  /**
   * Each substat's name and how many rarities it comes in, in save-file effect ID order.
   */
  substatClusters: [

    ["Attack Speed", 6],
    ["Critical Chance", 6],
    ["Critical Factor", 6],
    ["Attack Range", 6],
    ["Damage / Meter", 6],
    ["Multishot Chance", 5],
    ["Multishot Targets", 4],
    ["Rapid Fire Chance", 5],
    ["Rapid Fire Duration", 5],
    ["Bounce Shot Chance", 5],
    ["Bounce Shot Targets", 4],
    ["Bounce Shot Range", 5],
    ["Super Crit Chance", 4],
    ["Super Crit Multi", 4],
    ["Rend Armor Chance", 3],
    ["Rend Armor Multi", 3],
    ["Max Rend Armor Multi", 3],

    ["Health Regen", 6],
    ["Defense %", 6],
    ["Defense Absolute", 6],
    ["Thorns Damage", 4],
    ["Lifesteal", 4],
    ["Knockback Chance", 4],
    ["Knockback Force", 4],
    ["Orb Speed", 4],
    ["Orbs", 2],
    ["Shockwave Size", 4],
    ["Shockwave Frequency", 4],
    ["Land Mine Chance", 5],
    ["Land Mine Damage", 5],
    ["Land Mine Radius", 5],
    ["Death Defy", 3],
    ["Wall Health", 4],
    ["Wall Rebuild", 4],

    ["Cash Bonus", 6],
    ["Cash / Wave", 6],
    ["Coins / Kill Bonus", 6],
    ["Coins / Wave", 6],
    ["Free Attack Upgrade", 6],
    ["Free Defense Upgrade", 6],
    ["Free Utility Upgrade", 6],
    ["Interest / Wave", 4],
    ["Recovery Amount", 4],
    ["Package Chance", 4],
    ["Enemy Attack Level Skip", 4],
    ["Enemy Health Level Skip", 4],

    ["Chain Lightning - Damage", 6],
    ["Chain Lightning - Quantity", 4],
    ["Chain Lightning - Chance", 6],
    ["Smart Missiles - Damage", 6],
    ["Smart Missiles - Quantity", 4],
    ["Smart Missiles - Cooldown", 3],
    ["Death Wave - Damage", 6],
    ["Death Wave - Quantity", 3],
    ["Death Wave - Cooldown", 3],
    ["Chrono Field - Duration", 3],
    ["Chrono Field - Speed Reduction", 4],
    ["Chrono Field - Cooldown", 3],
    ["Inner Land Mines - Damage", 6],
    ["Inner Land Mines - Quantity", 3],
    ["Inner Land Mines - Cooldown", 4],
    ["Golden Tower - Bonus", 4],
    ["Golden Tower - Duration", 3],
    ["Golden Tower - Cooldown", 3],
    ["Poison Swamp - Damage", 6],
    ["Poison Swamp - Duration", 3],
    ["Poison Swamp - Cooldown", 5],
    ["Black Hole - Size", 6],
    ["Black Hole - Duration", 3],
    ["Black Hole - Cooldown", 3],
    ["Spotlight - Bonus", 6],
    ["Spotlight - Angle", 4],
    ["Spotlight - Quantity", 1],

    ["Max Recovery", 4],
  ],
};
