const ePathsReader = {

  /**
   * Reads EPaths data from a v5.09.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_09_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_09_00_00");

      var eHPRange = "eHP!AJ1:AY50";
      var eDamageRange = "eDamage!AI1:AY100";
      var eEconRange = "eEcon!AK1:AZ65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_09_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_09_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = ePathsEEcon.getVersion5_09_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_09_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.08.04.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_08_04_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_08_04_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AK1:AY65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = ePathsEEcon.getVersion5_08_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_08_04_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.08.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_08_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_08_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AS1:BG65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AQ3:AQ5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = ePathsEEcon.getVersion5_08_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_08_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.06.02.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_06_02_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_06_02_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;

      var eHPData = ePathsEHP.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = ePathsEEcon.getVersion5_06_02_00eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_06_02_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.05.01.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_05_01_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_05_01_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_05_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = ePathsEEcon.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_05_01_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.05.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_05_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_05_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_03_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion5_05_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = ePathsEEcon.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_05_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.03.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_03_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_03_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion5_03_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = ePathsEEcon.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_03_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.00.01.04 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_00_01_04: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version5_00_01_04");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX90";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AA3:AA5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion4_11_03_21eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = ePathsEEcon.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version5_00_01_04`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v4.11.03.21 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_11_03_21: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version4_11_03_21");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX75";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AA3:AA5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = ePathsEHP.getVersion4_11_03_21eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = ePathsEEcon.getVersion4_11_03_21eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version4_11_03_21`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v4.11.02.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_11_02_00: function (oldSheetID) {
    try {
      console.log("Called: ePathsReader.version4_11_02_00");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX75";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L4:L5";
      var eRegenLabRange = "eHP!AA4:AA5";
      var eDamageLabRange = "eDamage!L4:L5";
      var eEconLabRange = "eEcon!L4:L5";
      var eDiscountLabRange = "eEcon!AG4:AG5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;

      var eHPData = ePathsEHP.getVersion4_11_02_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = ePathsEDamage.getVersion4_11_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = ePathsEEcon.getVersion4_11_02_00eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsReader.version4_11_02_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },
};
