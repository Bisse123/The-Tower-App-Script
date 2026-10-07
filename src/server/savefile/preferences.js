const SAVE_FILE_WAVE_CAP_PREF_KEY = "SAVE_FILE_PLAYER_WAVES_AT_CAP";

/**
 * Client-callable. Whether the user treats player waves as capped.
 * @returns {boolean}
 */
function getSaveFilePlayerWaveCapPreference() {
  const pref = PropertiesService.getUserProperties().getProperty(
    SAVE_FILE_WAVE_CAP_PREF_KEY,
  );
  return pref === null ? true : pref === "true";
}

/**
 * Client-callable. Stores the wave-cap preference.
 * @param {boolean} atCap
 * @returns {void}
 */
function setSaveFilePlayerWaveCapPreference(atCap) {
  PropertiesService.getUserProperties().setProperty(
    SAVE_FILE_WAVE_CAP_PREF_KEY,
    String(!!atCap),
  );
  return true;
}
