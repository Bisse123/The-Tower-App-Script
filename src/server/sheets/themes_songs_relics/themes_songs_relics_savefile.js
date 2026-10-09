const themesAndRelicsHeaders = {
  towerSkins: "towerUnlocked",
  backgroundSkins: "backgroundUnlocked",
  menuSkins: "menuUnlocked",
  guardianSkins: "guardianSkinUnlocked",
  profileBanners: "profileBannerUnlocked",
  songs: "trackAvailable",
  relicsUnlocked: "relicsUnlocked",
}

const themesAndRelicsSaveFile = {

  /**
   * Parses Themes_Songs_Relics data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseThemesAndRelicsData: function (data) {
    try {
      var towerSkins = themesAndRelicsCatalog.towerSkins;
      var milestoneSkins = themesAndRelicsCatalog.milestoneSkins;
      var backgroundSkins = themesAndRelicsCatalog.backgroundSkins;
      var guardianSkins = themesAndRelicsCatalog.guardianSkins;
      var profileBanners = themesAndRelicsCatalog.profileBanners;
      var menuThemes = themesAndRelicsCatalog.menuThemes;
      var songs = themesAndRelicsCatalog.songs;
      var relics = themesAndRelicsCatalog.relics;

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
      var errorReport = errors.report("themesAndRelicsSaveFile.parseThemesAndRelicsData", error, {
        data: data,
        oldThemesNames: oldThemesNames,
        oldRelics: oldRelics,
      });
      return errors.fail(errorReport);
    }
  },
};
