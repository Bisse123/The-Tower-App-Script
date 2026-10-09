const sheetVars = (sheetType) => {
  var sheetTypeFunctions = {
    Laboratory: lab,
    Workshop: workshop,
    "Ultimate Weapon": ultimate,
    "Themes, Songs & Relics": themesAndRelics,
    "Themes & Songs": themes,
    Bots: bots,
    Relics: relics,
    Vault: vault,
    Cards: cards,
    Modules: modules,
    Guardians: guardians,
    "Player & Stuff": playerStuff,
    "IDS Collection": collection,
    "IDS Master": master,
    "Effective Paths": ePaths,
  };
  return sheetTypeFunctions[sheetType];
};

const spreadsheets = (spreadsheetTypeName, sheetID) => {
  if (!spreadsheetTypeName) {
    console.log(`No spreadsheet type name provided.`);
    return null;
  }

  const result = CacheManager.getSpreadsheet(spreadsheetTypeName, sheetID);

  if (!result) {
    if (!sheetID) {
      console.log(
        `Spreadsheet not found in cache and no sheet ID provided for: ${spreadsheetTypeName}`
      );
    } else {
      console.log(`Spreadsheet not found with ID: ${sheetID}`);
    }
    return null;
  }

  return result;
};
