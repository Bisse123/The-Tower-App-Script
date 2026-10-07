const CLIENT_ERROR_THROTTLE_SECONDS = 300;

/**
 * Snapshots a data map that arrived from a page, per key, sharing one budget.
 * @param {*} data
 * @returns {Object|null} Null when data is not a plain object.
 */
function _boundInboundData(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) return null;
  var budget = errors.budget();
  var out = {};
  Object.keys(data)
    .slice(0, 25)
    .forEach(function (key) {
      out[key] = errors.snapshot(data[key], 0, [], budget);
    });
  return out;
}

/**
 * Per-user throttle keyed on a fingerprint, holding the first caller's
 * reference for CLIENT_ERROR_THROTTLE_SECONDS.
 * @param {string} fingerprintKey
 * @param {string} reference Kept when this call is the first.
 * @returns {string} The already-written entry's reference, or "" to write.
 */
function _throttleReference(fingerprintKey, reference) {
  try {
    var key = Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, fingerprintKey),
    );
    var cache = CacheService.getUserCache();
    var existing = cache.get(key);
    if (existing) return existing;
    cache.put(key, reference || "1", CLIENT_ERROR_THROTTLE_SECONDS);
  } catch (ignored) {
  }
  return "";
}

/**
 * Client-callable. Logs a browser-side failure under the client service.
 * @param {{source?: string, message?: string, stack?: string, reference?: string,
 *   page?: string, viewType?: string, userAgent?: string, context?: Object}} payload
 * @returns {{success: boolean, reference: string, throttled?: boolean}}
 */
function reportClientError(payload) {
  try {
    var safe = payload && typeof payload === "object" ? payload : {};
    var source = String(safe.source || "client").slice(0, 120);
    var message = String(safe.message || "Unknown client error").slice(0, 2000);
    var stack = safe.stack ? String(safe.stack).slice(0, 8000) : "";
    var reference = String(safe.reference || errors.reference()).slice(0, 40);

    var alreadyLogged = _throttleReference(`${source}|${message}`, reference);
    if (alreadyLogged) {
      return {
        success: true,
        reference: alreadyLogged,
        throttled: true,
      };
    }

    console.error({
      "@type": ERROR_REPORT_TYPE,
      message: stack || `Error: ${message}\n    at ${source} (${source}:0:0)`,
      serviceContext: {
        service: "the-tower-app-script-client",
        version: errors._runningVersion(),
      },
      context: {
        reportLocation: { functionName: source },
        user: errors.userKey(),
      },
      reference: reference,
      source: source,
      code: errors.CODES.CLIENT,
      kind: "bug",
      detail: message,
      data: {
        page: String(safe.page || "").slice(0, 120),
        viewType: String(safe.viewType || "").slice(0, 40),
        userAgent: String(safe.userAgent || "").slice(0, 300),
        context: errors.snapshot(safe.context || {}),
      },
    });

    return { success: true, reference: reference };
  } catch (error) {
    console.error(`reportClientError failed: ${error && error.message}`);
    return { success: false, reference: "" };
  }
}

/**
 * Client-callable. Writes the deferred entry for a server bug the browser has
 * finished receiving, as kind "bug".
 * @param {{trace?: string[], code?: string, detail?: string, message?: string,
 *   reference?: string, stack?: string, note?: string, data?: Object}} payload
 * @returns {{success: boolean, reference: string, throttled?: boolean}}
 */
function reportServerError(payload) {
  try {
    var safe = payload && typeof payload === "object" ? payload : {};
    var trace =
      Array.isArray(safe.trace) && safe.trace.length
        ? safe.trace.slice(0, 20).map(function (s) {
            return String(s).slice(0, 120);
          })
        : ["unknown"];
    var source = trace[0];
    var code = String(safe.code || errors.CODES.INTERNAL).slice(0, 40);
    var detail = String(safe.detail || safe.message || "").slice(0, 4000);
    var reference = String(safe.reference || errors.reference()).slice(0, 40);
    var stack = safe.stack ? String(safe.stack).slice(0, 8000) : "";

    var alreadyLogged = _throttleReference(`${source}|${detail}`, reference);
    if (alreadyLogged) {
      return {
        success: true,
        reference: alreadyLogged,
        throttled: true,
      };
    }

    var entry = errors._event({
      service: "the-tower-app-script",
      source: source,
      code: code,
      reference: reference,
      kind: "bug",
      detail: detail,
      trace: trace,
      stack: stack,
      note: safe.note ? String(safe.note).slice(0, 500) : "",
      data: _boundInboundData(safe.data),
    });

    errors._write(entry);
    return { success: true, reference: reference };
  } catch (error) {
    console.error(`reportServerError failed: ${error && error.message}`);
    return { success: false, reference: "" };
  }
}

/**
 * Client-callable, and called from 22_error_scripts.html's scriptlet. The codes
 * and expected flags the browser needs, as inlinable JSON.
 * @returns {string} A JSON object literal, e.g. '{"CODES":{…},"MESSAGES":{…},"EXPECTED":{…}}'.
 */
function errorContract() {
  const contract = errors.contract();
  return contract;
}
