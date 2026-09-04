const ERROR_REPORT_TYPE =
  "type.googleapis.com/google.devtools.clouderrorreporting.v1beta1.ReportedErrorEvent";

const CLIENT_ERROR_THROTTLE_SECONDS = 300;

const SNAPSHOT_MAX_DEPTH = 6;
const SNAPSHOT_MAX_NODES = 5000;
const SNAPSHOT_MAX_CHARS = 100000;

const errors = {
  CODES: {
    ACCESS_DENIED: "ACCESS_DENIED",
    NOT_FOUND: "NOT_FOUND",
    INVALID_INPUT: "INVALID_INPUT",
    INVALID_LINK: "INVALID_LINK",
    INVALID_FILE: "INVALID_FILE",
    SHEET_STRUCTURE: "SHEET_STRUCTURE",
    VERSION_OUTDATED: "VERSION_OUTDATED",
    QUOTA: "QUOTA",
    TIMEOUT: "TIMEOUT",
    CLIENT: "CLIENT",
    INTERNAL: "INTERNAL",
    RECOVERED: "RECOVERED",
  },

  MESSAGES: {
    ACCESS_DENIED:
      "The script does not have access to that file. Grant access and try again.",
    NOT_FOUND: "That file could not be found. It may have been moved or deleted.",
    INVALID_INPUT: "Something was missing from that request. Please reload and try again.",
    INVALID_LINK:
      "That does not look like a Google Sheets™ link or ID. Check it and try again.",
    INVALID_FILE:
      "That does not look like a playerInfo.dat save file. Check you picked the right file and try again.",
    SHEET_STRUCTURE:
      "The script could not find something it needs inside your sheet.",
    VERSION_OUTDATED: "That sheet is not a version this step can work with.",
    QUOTA:
      "Google is rate-limiting this account right now. Wait a few minutes and try again.",
    TIMEOUT: "That took too long to finish. Try again with fewer sheets at once.",
    CLIENT: "Something went wrong in this page.",
    INTERNAL: "Something went wrong on our side.",
    RECOVERED: "Something did not work, but the script carried on without it.",
  },

  EXPECTED: {
    ACCESS_DENIED: true,
    INVALID_LINK: true,
    INVALID_FILE: true,
    NOT_FOUND: true,
    QUOTA: true,
    TIMEOUT: true,
    VERSION_OUTDATED: true,
    RECOVERED: true,
  },

  /**
   * Whether a code is an expected outcome rather than a defect.
   * @param {string} code One of errors.CODES.
   * @returns {boolean}
   */
  isExpected: function (code) {
    return errors.EXPECTED[code] === true;
  },

  /**
   * Copies every own key of source onto target, in place.
   * @param {Object} target
   * @param {Object} source
   * @returns {Object} target
   */
  merge: function (target, source) {
    Object.keys(source).forEach(function (key) {
      target[key] = source[key];
    });
    return target;
  },

  /**
   * A fresh snapshot allowance for one log entry.
   * @returns {{nodes: number, chars: number}}
   */
  budget: function () {
    return { nodes: SNAPSHOT_MAX_NODES, chars: SNAPSHOT_MAX_CHARS };
  },

  /**
   * Copies a value into a form safe to log, capped in depth, breadth, string
   * length and total size. Never throws.
   * @param {*} value
   * @param {number} [depth] Current recursion level; omit at the call site.
   * @param {Array} [seen] Ancestors, for circular detection; omit at the call site.
   * @param {{nodes: number, chars: number}} [budget] Share one across an entry.
   * @returns {*} A scalar, array or plain object, or a "[…]" marker string.
   */
  snapshot: function (value, depth, seen, budget) {
    depth = depth || 0;
    seen = seen || [];
    budget = budget || errors.budget();
    try {
      budget.nodes--;

      if (value === undefined) return "[undefined]";
      if (value === null) return null;
      if (typeof value === "function") return "[Function]";
      if (typeof value === "string") {
        var text =
          value.length > 300
            ? value.slice(0, 300) + `… (${value.length} chars)`
            : value;
        budget.chars -= text.length + 2;
        return text;
      }
      if (typeof value === "number" || typeof value === "boolean") {
        budget.chars -= 8;
        return value;
      }
      if (typeof value !== "object") {
        var other = String(value);
        budget.chars -= other.length + 2;
        return other;
      }
      if (seen.indexOf(value) !== -1) return "[Circular]";
      if (budget.nodes <= 0 || budget.chars <= 0) {
        return "[truncated: entry size]";
      }
      var nextSeen = seen.concat([value]);
      budget.chars -= 2;

      if (Array.isArray(value)) {
        if (depth >= SNAPSHOT_MAX_DEPTH) return `[Array(${value.length})]`;
        var limit = 10;
        var items = value.slice(0, limit).map(function (item) {
          return errors.snapshot(item, depth + 1, nextSeen, budget);
        });
        if (value.length > limit) {
          items.push(`… (${value.length - limit} more of ${value.length})`);
        }
        return items;
      }

      if (depth >= SNAPSHOT_MAX_DEPTH) return "[Object]";
      var keys = Object.keys(value);
      var out = {};
      keys.slice(0, 25).forEach(function (key) {
        budget.chars -= String(key).length + 4;
        out[key] = errors.snapshot(value[key], depth + 1, nextSeen, budget);
      });
      if (keys.length > 25) {
        out.__more__ = `${keys.length - 25} more key(s)`;
      }
      return out;
    } catch (ignored) {
      return "[unserializable]";
    }
  },

  _lastReference: "",
  _lastRecordedAt: 0,
  _lastCode: "",
  _lastSource: "",
  _lastDetail: "",
  _lastNote: "",
  _lastData: null,
  _lastStack: "",

  /**
   * Mints a short quotable id, e.g. "TWR-M4X2K9-A7F3".
   * @returns {string}
   */
  reference: function () {
    var stamp = Date.now().toString(36).toUpperCase();
    var salt = Math.floor(Math.random() * 0xffff)
      .toString(36)
      .toUpperCase();
    return `TWR-${stamp}-${salt}`;
  },

  /**
   * Maps a thrown value to a code by matching Google's wording.
   * @param {*} error
   * @returns {string} One of errors.CODES; INTERNAL when nothing matches.
   */
  classify: function (error) {
    var text = errors.text(error).toLowerCase();

    if (!text) return errors.CODES.INTERNAL;

    if (
      text.indexOf("unable to parse range") !== -1 ||
      text.indexOf("invalid range") !== -1 ||
      text.indexOf("range not found") !== -1 ||
      text.indexOf("no grid with id") !== -1
    ) {
      return errors.CODES.SHEET_STRUCTURE;
    }
    if (
      text.indexOf("permission") !== -1 ||
      text.indexOf("do not have access") !== -1 ||
      text.indexOf("does not have permission") !== -1 ||
      text.indexOf("access denied") !== -1 ||
      text.indexOf("unauthorized") !== -1 ||
      text.indexOf("forbidden") !== -1
    ) {
      return errors.CODES.ACCESS_DENIED;
    }

    if (
      text.indexOf("file not found") !== -1 ||
      text.indexOf("no item with the given id") !== -1 ||
      text.indexOf("requested entity was not found") !== -1
    ) {
      return errors.CODES.NOT_FOUND;
    }
    if (
      text.indexOf("quota") !== -1 ||
      text.indexOf("rate limit") !== -1 ||
      text.indexOf("too many requests") !== -1 ||
      text.indexOf("user rate limit exceeded") !== -1 ||
      text.indexOf("service invoked too many times") !== -1
    ) {
      return errors.CODES.QUOTA;
    }
    if (
      text.indexOf("timed out") !== -1 ||
      text.indexOf("timeout") !== -1 ||
      text.indexOf("exceeded maximum execution time") !== -1
    ) {
      return errors.CODES.TIMEOUT;
    }
    return errors.CODES.INTERNAL;
  },

  /**
   * Whatever was thrown, as text. Empty string when there is nothing.
   * @param {*} error
   * @returns {string}
   */
  text: function (error) {
    if (!error) return "";
    if (typeof error === "string") return error;

    if (typeof error.message === "string") return error.message;
    if (error.message) return String(error.message);
    try {
      return String(error);
    } catch (ignored) {
      return "Unstringifiable error";
    }
  },

  /**
   * The thrown value's stack, or a synthetic one-frame stack naming source.
   * The first line is never empty.
   * @param {string} source
   * @param {*} [error]
   * @param {string} [summary] Stands in when error carries no text.
   * @returns {string}
   */
  stack: function (source, error, summary) {
    if (error && error.stack) return String(error.stack);
    var name = error && error.name ? error.name : "Error";
    var text = errors.text(error) || summary || source;
    return `${name}: ${text}\n    at ${source} (${source}:0:0)`;
  },

  /**
   * The active user as 12 hex characters of an MD5 of their address.
   * @returns {string} The hash, or "anonymous" / "unknown".
   */
  userKey: function () {
    try {
      var email = Session.getActiveUser().getEmail();
      if (!email) return "anonymous";
      var digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.MD5,
        email,
      );
      return digest
        .map(function (byte) {
          return ((byte & 0xff) + 0x100).toString(16).slice(1);
        })
        .join("")
        .slice(0, 12);
    } catch (ignored) {
      return "unknown";
    }
  },

  /**
   * Adds the outdated-release notice to a failure envelope, when there is one.
   * @param {Object} envelope Mutated in place.
   * @returns {Object} The same envelope.
   */
  /**
   * The running release, for serviceContext. Tolerates appVersion being
   * absent so a log entry is never lost over a missing version label.
   * @returns {string}
   */
  _runningVersion: function () {
    try {
      return appVersion.running();
    } catch (ignored) {
      return "unversioned";
    }
  },

  _attachOutdated: function (envelope) {
    try {
      var stale = appVersion.status();
      if (stale) envelope.outdated = stale;
    } catch (ignored) {
    }
    return envelope;
  },

  /**
   * Builds the ReportedErrorEvent payload both bug writers emit.
   * @param {{service: string, source: string, code: string, kind: string,
   *   detail: string, trace: string[], reference?: string, stack?: string,
   *   note?: string, data?: Object}} fields
   * @returns {Object} A log entry; reference, note and data are omitted when empty.
   */
  _event: function (fields) {
    var detail = fields.detail || "";
    var entry = {
      "@type": ERROR_REPORT_TYPE,
      message:
        fields.stack ||
        errors.stack(
          fields.source,
          { message: detail },
          `${fields.code} in ${fields.source}`,
        ),
      serviceContext: {
        service: fields.service,
        version: errors._runningVersion(),
      },
      context: {
        reportLocation: { functionName: fields.source },
        user: errors.userKey(),
      },
      source: fields.source,
      code: fields.code,
      expected: false,
      kind: fields.kind,
      detail: detail,
      trace: fields.trace,
    };

    if (fields.reference) entry.reference = fields.reference;
    if (fields.note) entry.note = fields.note;
    if (fields.data) entry.data = fields.data;
    return entry;
  },

  /**
   * Records a caught exception. Classifies it unless code is given.
   * @param {string} source
   * @param {*} error
   * @param {Object} [context] Diagnostics; values are snapshotted.
   * @param {string} [code] Overrides classification.
   * @returns {Object} { reference, code, expected, message, detail, source,
   *   trace, note?, data?, stack? }
   */
  report: function (source, error, context, code) {
    var resolvedCode = code || errors.classify(error);
    var detail = errors.text(error);
    return errors.record(source, resolvedCode, detail, error, context);
  },

  /**
   * Builds a report and, for an expected outcome only, writes it at WARNING.
   * A bug's write is deferred. Never throws.
   * @param {string} source
   * @param {string} code
   * @param {string} detail
   * @param {*} [error]
   * @param {Object} [context] "note" is reserved; other keys become data.
   * @returns {Object} The report; reference is empty when expected.
   */
  record: function (source, code, detail, error, context) {
    var expected = errors.isExpected(code);
    var reference = expected ? "" : errors.reference();

    var note = "";
    var data = null;
    if (context) {
      var budget = errors.budget();
      Object.keys(context).forEach(function (key) {
        if (key === "note") {
          if (context.note) note = context.note;
          return;
        }
        data = data || {};
        data[key] = errors.snapshot(context[key], 0, [], budget);
      });
    }

    var stack = error && error.stack ? String(error.stack) : "";

    errors._lastReference = reference;
    errors._lastRecordedAt = Date.now();
    errors._lastCode = code;
    errors._lastSource = source;
    errors._lastDetail = detail;
    errors._lastNote = note;
    errors._lastData = data;
    errors._lastStack = stack;

    if (expected) {
      var fields = {
        source: source,
        code: code,
        expected: expected,
        kind: "expected",
      };
      if (detail) fields.detail = detail;
      if (note) fields.note = note;
      if (data) fields.data = data;
      errors._write(
        errors.merge({ message: `${source}: ${detail || code}` }, fields),
      );
    }

    var report = {
      reference: reference,
      code: code,
      expected: expected,
      message: errors.MESSAGES[code] || errors.MESSAGES.INTERNAL,
      detail: detail,
      source: source,
      trace: [source],
    };
    if (note) report.note = note;
    if (data) report.data = data;
    if (stack) report.stack = stack;
    return report;
  },

  /**
   * The only console call in the module. WARNING when expected, else ERROR.
   * @param {Object} entry
   * @returns {void}
   */
  _write: function (entry) {
    try {
      if (entry.expected) {
        console.warn(entry);
      } else {
        console.error(entry);
      }
    } catch (loggingError) {
      console.error(`${entry.reference || entry.code} ${entry.source}: ${entry.detail}`);
    }
  },

  /**
   * Turns a report into the failure envelope the client reads.
   * @param {Object} report
   * @param {string} [message] Overrides the user-facing text.
   * @param {Object} [extra] Extra envelope fields.
   * @returns {{success: false, code: string, expected: boolean, message: string,
   *   reference: string, detail: string, trace: string[]}} Plus note, data and
   *   stack when the failure is a bug.
   */
  fail: function (report, message, extra) {
    var envelope = {
      success: false,
      code: (report && report.code) || errors.CODES.INTERNAL,
      expected: Boolean(report && report.expected),
      message: message || (report && report.message) || errors.MESSAGES.INTERNAL,
      reference: (report && report.reference) || "",
      detail: (report && report.detail) || "",
      trace: (report && report.trace) || (report && [report.source]) || [],
    };

    if (!envelope.expected) {
      if (report && report.note) envelope.note = report.note;
      if (report && report.data) envelope.data = report.data;
      if (report && report.stack) envelope.stack = report.stack;
    }

    errors._attachOutdated(envelope);

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },

  /**
   * Relays an inner failure outward without recording a second incident,
   * extending the trace. Falls through to reject when nothing was recorded.
   * @param {string} source
   * @param {*} inner Whatever the inner call returned.
   * @param {string} [message] Overrides the user-facing text.
   * @param {Object} [extra] Extra envelope fields.
   * @returns {Object} A failure envelope, same shape as fail().
   */
  propagate: function (source, inner, message, extra) {
    var fresh = Date.now() - errors._lastRecordedAt < 60000;

    var code =
      (inner && inner.code) ||
      (fresh ? errors._lastCode : "") ||
      errors.CODES.INTERNAL;
    // RECOVERED says the code caught something and carried on. It describes
    // that earlier moment, not this failure, so it is never what a caller is
    // reporting - its diagnostics below are still worth borrowing.
    if (code === errors.CODES.RECOVERED) code = errors.CODES.INTERNAL;
    var reference =
      (inner && inner.reference) || (fresh ? errors._lastReference : "");

    var detail =
      (inner && inner.detail) || (fresh ? errors._lastDetail : "") || "";
    var trace = (inner && inner.trace) ||
      (fresh && errors._lastSource ? [errors._lastSource] : []);
    var note = (inner && inner.note) || (fresh ? errors._lastNote : "") || "";
    var data = (inner && inner.data) || (fresh ? errors._lastData : null);
    var stack = (inner && inner.stack) || (fresh ? errors._lastStack : "") || "";

    if (!(inner && inner.code) && !fresh) {
      return errors.reject(source, code, message || (inner && inner.message), extra, {
        note: inner && inner.message,
      });
    }

    var envelope = {
      success: false,
      code: code,
      expected:
        inner && typeof inner.expected === "boolean"
          ? inner.expected
          : errors.isExpected(code),
      message:
        !(inner && inner.code) && errors.isExpected(code)
          ? errors.MESSAGES[code]
          : message || (inner && inner.message) || errors.MESSAGES[code],
      reference: reference,
      detail: detail,
      trace: trace.concat(source),
    };

    if (!envelope.expected) {
      if (note) envelope.note = note;
      if (data) envelope.data = data;
      if (stack) envelope.stack = stack;
    }

    errors._attachOutdated(envelope);

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },

  /**
   * Records a failure the code diagnosed itself, with no exception behind it.
   * @param {string} source
   * @param {string} code
   * @param {string} message What the user is told.
   * @param {Object} [extra] Extra envelope fields.
   * @param {Object} [context] Diagnostics; stays server-side when expected.
   * @returns {Object} A failure envelope, same shape as fail().
   */
  reject: function (source, code, message, extra, context) {
    var resolved = message || errors.MESSAGES[code] || errors.MESSAGES.INTERNAL;
    var report = errors.record(source, code, resolved, null, context);

    var envelope = {
      success: false,
      code: code,
      expected: report.expected,
      message: resolved,
      reference: report.reference,
      detail: report.expected ? "" : report.note || "",
      trace: report.trace,
    };
    if (!envelope.expected) {
      if (report.note) envelope.note = report.note;
      if (report.data) envelope.data = report.data;
    }

    errors._attachOutdated(envelope);

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },
};

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
