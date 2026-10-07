/**
 * Web app entry point. Serves the save-file, Get Started or update page
 * according to the query parameters.
 * @param {{parameter?: Object}} e
 * @returns {HtmlOutput}
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  var targetPage = (params.page || "").toString().toLowerCase();

  if (targetPage === "savefile" || targetPage === "save-file" || params.saveFile === "true") {
    var sheetType = params.sheetType || "";
    if (sheetType === "IDS Collection - all IDS-Sheets on one file") {
      sheetType = "IDS Collection";
    }
    var saveFileTemplate = HtmlService.createTemplateFromFile("20_SavedFileApp");
    saveFileTemplate.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    saveFileTemplate.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");
    saveFileTemplate.viewType = "webapp";
    saveFileTemplate.idMasterID = params.idMasterID
      ? sheetRefs.extractSheetId(params.idMasterID) || ""
      : "";
    saveFileTemplate.sheetType = sheetType;
    return saveFileTemplate
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Import Data From Game");
  }

  var openGetStarted =
    targetPage === "getstarted" ||
    targetPage === "get-started" ||
    params.getStarted === "true";

  if (openGetStarted) {
    var getStartedTemplate = HtmlService.createTemplateFromFile("20_getStartedApp");
    getStartedTemplate.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    getStartedTemplate.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");
    var effectivePathsID =
      sheetRefs.extractSheetId(params.effectivePathsID) || "";
    getStartedTemplate.effectivePathsID = effectivePathsID;
    getStartedTemplate.viewType = "webapp";
    return getStartedTemplate
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Get Started");
  }

  var template = HtmlService.createTemplateFromFile("client/pages/update");
  var newSheetID = sheetRefs.extractSheetId(params.newSheetID) || "";
  var oldSheetID = sheetRefs.extractSheetId(params.oldSheetID) || "";
  var idMasterID = sheetRefs.extractSheetId(params.idMasterID) || "";
  var sheetType = params.sheetType || "";

  if (sheetType === "IDS Master") {
    if (!oldSheetID && idMasterID) {
      oldSheetID = idMasterID;
    }

    console.log("IDS Master detected in doGet, setting parameters accordingly");
    template.newSheetID = "";
    template.oldSheetID = oldSheetID;
    template.idMasterID = "";
    template.sheetType = sheetType;
  } else {
    if (typeof sheetType === "string" && sheetType.includes("IDS Collection")) {
      sheetType = "IDS Collection";
    }
    template.newSheetID = newSheetID;
    template.oldSheetID = oldSheetID;
    template.idMasterID = idMasterID;
    template.sheetType = sheetType;
  }

  template.API_KEY =
    PropertiesService.getScriptProperties().getProperty("API_KEY");
  template.APP_ID =
    PropertiesService.getScriptProperties().getProperty("APP_ID");

  template.viewType = "webapp";
  template.accessRequired = false;

  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setTitle("Import Data");
}

/**
 * Inlines another HTML file, for the <?!= include(...) ?> templating.
 * @param {string} filename
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Inlines another HTML file, evaluating its scriptlets first, for the
 * <?!= includeTemplate(...) ?> templating.
 * @param {string} filename
 * @returns {string}
 */
function includeTemplate(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
