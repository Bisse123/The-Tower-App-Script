const appVersion = {
  /**
   * The release this code belongs to, and the oldest release still supported.
   * Both are rewritten by `npm run bump` — do not edit them by hand.
   *
   * They are baked into the source rather than read from script properties
   * because properties are project-scoped: every version and every deployment
   * of this project shares one store, so a property would report the newest
   * release to everybody, including a user still running an older add-on. A
   * constant ships with the code that contains it, so it always names what is
   * actually running.
   *
   * MINIMUM is this release's declaration of the floor, and it has to live
   * here rather than in package.json: clasp pushes only src/, so the running
   * script never sees package.json.
   */
  VERSION: "5.0.0",
  MINIMUM: "5.0.0",

  /**
   * The release this code was built from.
   * @returns {string} "unversioned" when nothing was baked in.
   */
  running: function () {
    return appVersion.VERSION || "unversioned";
  },

  /**
   * The newest release published so far, from a script property. Shared across
   * every version and deployment, which is what makes it usable as a pointer.
   * @returns {string} "" when nothing has been published yet.
   */
  latest: function () {
    try {
      return (
        PropertiesService.getScriptProperties().getProperty(
          "LATEST_APP_VERSION",
        ) || ""
      );
    } catch (ignored) {
      return "";
    }
  },

  /**
   * The oldest release still considered usable, from a script property.
   * @returns {string} "" when no floor has been published.
   */
  minimum: function () {
    try {
      return (
        PropertiesService.getScriptProperties().getProperty(
          "MINIMUM_APP_VERSION",
        ) || ""
      );
    } catch (ignored) {
      return "";
    }
  },

  /**
   * Numeric version comparison, part by part.
   * @param {string} a
   * @param {string} b
   * @returns {boolean} True when a is strictly older than b.
   */
  isOlder: function (a, b) {
    var parse = function (v) {
      var match = String(v || "").match(/\d+(?:\.\d+)*/);
      return match ? match[0].split(".").map(Number) : [];
    };
    var left = parse(a);
    var right = parse(b);
    if (!right.length) return false;
    if (!left.length) return true;
    for (var i = 0; i < Math.max(left.length, right.length); i++) {
      var l = left[i] || 0;
      var r = right[i] || 0;
      if (l < r) return true;
      if (l > r) return false;
    }
    return false;
  },

  /**
   * Whether this copy is behind the newest release, and whether it is below
   * the supported floor. Never throws.
   * @returns {{running: string, latest: string, minimum: string,
   *   unsupported: boolean}|null} Null when it is current.
   */
  status: function () {
    try {
      var running = appVersion.running();
      if (running === "unversioned") return null;
      var latest = appVersion.latest();
      var minimum = appVersion.minimum();
      var behind = latest ? appVersion.isOlder(running, latest) : false;
      var unsupported = minimum ? appVersion.isOlder(running, minimum) : false;
      if (!behind && !unsupported) return null;
      return {
        running: running,
        latest: latest,
        minimum: minimum,
        unsupported: unsupported,
      };
    } catch (ignored) {
      return null;
    }
  },

  /**
   * Records this code's version as the newest release, and the floor it
   * declares. Every older copy reads those properties to find out it is
   * behind, so nothing but a real release should call this.
   *
   * Throws rather than returning a failure: `clasp run` reports the return
   * value but exits 0 either way, so only an exception makes a bad release go
   * red in the deploy workflow.
   * @returns {{success: true, version: string, previous: string,
   *   minimum: string, previousMinimum: string}}
   * @throws {Error} When there is no baked version to publish.
   */
  publish: function () {
    var running = appVersion.running();
    if (running === "unversioned") {
      throw new Error(
        "No version is baked into this build, so there is nothing to " +
          "publish. Run `npm run bump` and redeploy.",
      );
    }

    var properties = PropertiesService.getScriptProperties();
    var previousLatest = properties.getProperty("LATEST_APP_VERSION") || "";
    properties.setProperty("LATEST_APP_VERSION", running);

    var minimum = appVersion.MINIMUM || "";
    var previousMinimum = properties.getProperty("MINIMUM_APP_VERSION") || "";
    if (minimum) {
      properties.setProperty("MINIMUM_APP_VERSION", minimum);
    }

    return {
      success: true,
      version: running,
      previous: previousLatest,
      minimum: minimum,
      previousMinimum: previousMinimum,
    };
  },
};

/**
 * Client-callable. Whether this copy of the script is behind the newest
 * release, and whether it is below the supported floor.
 *
 * Top-level on purpose: google.script.run can only reach functions declared at
 * the top level, not members of an object.
 * @returns {{running: string, latest: string, minimum: string,
 *   unsupported: boolean}|null} Null when it is current.
 */
function getAppVersionStatus() {
  return appVersion.status();
}

/**
 * Called by the deploy workflow after publishing main, via `clasp run`, which
 * likewise only reaches top-level functions.
 * @returns {{success: true, version: string, previous: string,
 *   minimum: string, previousMinimum: string}}
 * @throws {Error} When there is no baked version to publish.
 */
function setLatestAppVersion() {
  return appVersion.publish();
}
