/**
 * Error reporting for the whole project.
 *
 * Every catch block in the backend funnels through `errors.report`, and every
 * precondition the code checks itself through `errors.reject`. Both land in
 * `errors.record`, which sorts the failure into one of exactly two kinds:
 *
 *   EXPECTED — the app working as designed on input it cannot accept. A sheet
 *   too old to convert, a file the user never granted, Google rate-limiting
 *   the account. Logged at WARNING, kept out of Error Reporting, and given no
 *   reference: there is nothing here for anyone to report.
 *
 *   BUG — a range that should have been there and was not, an unclassified
 *   exception, a browser crash. Eventually logged at ERROR with a stack in
 *   the shape Cloud Error Reporting ingests, and given a short reference the
 *   user can quote back to us.
 *
 * `errors.EXPECTED` is the whole of that decision, and the code alone decides,
 * so a quota exception caught in a try/catch is treated exactly like one the
 * code detected itself.
 *
 *     } catch (error) {
 *       var errorReport = errors.report("exportData", error, { sheetType: sheetType });
 *       return errors.fail(errorReport);
 *     }
 *
 * A bug's log entry is *not* written at that catch block. `record` cannot
 * tell whether this catch is the one `google.script.run` actually invoked or
 * just another frame in a longer relay — Apps Script gives a function no way
 * to ask that — and writing there anyway is how the same incident used to
 * show up as several separate rows in Logs Explorer, one per hop, none of
 * them naming the full path. Instead the failure rides the normal return
 * value out to the browser, gaining a frame at every `errors.propagate` hop
 * on the way, and the browser — which unambiguously *is* the end of the line
 * — calls back to `reportServerError` with everything it has once it sees a
 * bug it did not already know about. That is where the one entry for this
 * incident actually gets written, with the complete trace already on it.
 *
 * That leaves one case the round trip cannot cover: a failure the code
 * catches, logs and then carries on from, so that no envelope is ever
 * returned and no browser ever sees it — a cache that would not open, an
 * optional cleanup step, a dialog that failed before it could be shown.
 * Those use `errors.reportFinal`, which writes the entry on the spot,
 * because for them there is no round trip coming and never was.
 *
 * See documentation/08-error-handling.md.
 */

const ERROR_REPORT_TYPE =
  "type.googleapis.com/google.devtools.clouderrorreporting.v1beta1.ReportedErrorEvent";

const CLIENT_ERROR_THROTTLE_SECONDS = 300;

const errors = {
  /**
   * Error codes. The client switches on these instead of matching message
   * text, so renaming a message never changes behaviour.
   */
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
  },

  /**
   * Plain-language text per code. Written for the person looking at the
   * sidebar, never for whoever reads the logs.
   */
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
    // Deliberately says nothing about what to do: what a user can do about an
    // out-of-date sheet depends entirely on which workflow they are in, so the
    // call site says it. This is only the fallback.
    VERSION_OUTDATED: "That sheet is not a version this step can work with.",
    QUOTA:
      "Google is rate-limiting this account right now. Wait a few minutes and try again.",
    TIMEOUT: "That took too long to finish. Try again with fewer sheets at once.",
    CLIENT: "Something went wrong in this page.",
    INTERNAL: "Something went wrong on our side.",
  },

  /**
   * The line between the two kinds of failure.
   *
   * EXPECTED — the app working as designed on input it cannot accept: a sheet
   * too old to convert, a file the user never granted, Google rate-limiting
   * the account. There is no defect here. These are logged at WARNING, stay
   * out of Error Reporting, and carry **no reference**: asking someone to
   * report a working system wastes their time and ours.
   *
   * Everything else is a BUG — a range that should have been there and was
   * not, an unclassified exception, a browser crash. Those are logged at ERROR
   * with a stack, reach Error Reporting, and carry a reference for the user to
   * quote.
   *
   * The code alone decides, whichever entry point produced it, so a quota
   * exception caught in a try/catch is treated exactly like one the code
   * detected itself.
   */
  EXPECTED: {
    ACCESS_DENIED: true,
    INVALID_LINK: true,
    INVALID_FILE: true,
    NOT_FOUND: true,
    QUOTA: true,
    TIMEOUT: true,
    VERSION_OUTDATED: true,
  },

  isExpected: function (code) {
    return errors.EXPECTED[code] === true;
  },

  /** Object.assign, which Apps Script's V8 has but older habits do not. */
  merge: function (target, source) {
    Object.keys(source).forEach(function (key) {
      target[key] = source[key];
    });
    return target;
  },

  /**
   * Bounds any value down to something safe to put in a log entry, so raw
   * locals can be passed to `report`/`reject` context without a second
   * thought: a 2000-row sheet range comes back as its first 10 rows and a
   * count of how many were cut, not a multi-megabyte log entry. Depth, breadth
   * and string length are all capped, a repeated object short-circuits as
   * "[Circular]" instead of recursing into itself, and nothing in here can
   * throw - a value that resists serialising becomes "[unserializable]"
   * rather than losing the whole log entry over it.
   */
  snapshot: function (value, depth, seen) {
    depth = depth || 0;
    seen = seen || [];
    try {
      if (value === undefined) return "[undefined]";
      if (value === null) return null;
      if (typeof value === "function") return "[Function]";
      if (typeof value === "string") {
        return value.length > 300
          ? value.slice(0, 300) + `… (${value.length} chars)`
          : value;
      }
      if (typeof value === "number" || typeof value === "boolean") {
        return value;
      }
      if (typeof value !== "object") return String(value);
      if (seen.indexOf(value) !== -1) return "[Circular]";
      var nextSeen = seen.concat([value]);

      if (Array.isArray(value)) {
        if (depth >= 3) return `[Array(${value.length})]`;
        var limit = 10;
        var items = value.slice(0, limit).map(function (item) {
          return errors.snapshot(item, depth + 1, nextSeen);
        });
        if (value.length > limit) {
          items.push(`… (${value.length - limit} more of ${value.length})`);
        }
        return items;
      }

      if (depth >= 3) return "[Object]";
      var keys = Object.keys(value);
      var out = {};
      keys.slice(0, 25).forEach(function (key) {
        out[key] = errors.snapshot(value[key], depth + 1, nextSeen);
      });
      if (keys.length > 25) {
        out.__more__ = `${keys.length - 25} more key(s)`;
      }
      return out;
    } catch (ignored) {
      return "[unserializable]";
    }
  },

  /**
   * The most recent failure recorded in this execution, whichever kind it was.
   * Apps Script gives every execution a fresh global state, so this is
   * per-request; the timestamp is belt-and-braces in case a context is ever
   * reused. `propagate` reads these when an inner call recorded a failure and
   * then swallowed it — the reference is empty for an expected one, and that
   * emptiness is itself the information it needs.
   */
  _lastReference: "",
  _lastRecordedAt: 0,
  _lastCode: "",
  _lastSource: "",
  _lastDetail: "",
  _lastNote: "",
  _lastData: null,
  _lastStack: "",

  /**
   * Close the books on the last recorded failure so nothing downstream can
   * borrow it.
   *
   * `propagate` falls back to `_last*` when the value it is relaying carries
   * no code of its own — that is how a `SheetsAPI` wrapper that reported and
   * returned `null` still gets its reference onto the envelope. But a failure
   * the code *deliberately swallowed* (see `reportFinal`) is finished: it has
   * already been written, nothing is relaying it, and leaving it sitting in
   * `_last*` means the next unrelated `propagate` in the same execution
   * inherits its reference, trace, note and data — reporting, say, a
   * Laboratory version mismatch as a Cards parser crash.
   */
  _forget: function () {
    errors._lastReference = "";
    errors._lastRecordedAt = 0;
    errors._lastCode = "";
    errors._lastSource = "";
    errors._lastDetail = "";
    errors._lastNote = "";
    errors._lastData = null;
    errors._lastStack = "";
  },

  /**
   * Short, human-quotable id, issued for bugs only. Ends up in the log entry
   * and in the panel, so a screenshot is enough to find the entry:
   *   jsonPayload.reference="TWR-M4X2K9-A7F3"
   */
  reference: function () {
    var stamp = Date.now().toString(36).toUpperCase();
    var salt = Math.floor(Math.random() * 0xffff)
      .toString(36)
      .toUpperCase();
    return `TWR-${stamp}-${salt}`;
  },

  /**
   * Best-effort mapping from a raw Google/Apps Script exception to one of our
   * codes, so a mechanically converted catch block still produces a message
   * worth reading.
   */
  classify: function (error) {
    var text = errors.text(error).toLowerCase();

    if (!text) return errors.CODES.INTERNAL;
    // Sheets tells these apart for us. A bad range comes back 400 "Unable to
    // parse range: <name>"; a quota comes back 429 "Quota exceeded …" or
    // "User rate limit exceeded". Same exception class, different text — so a
    // tab we asked for and did not get is a defect, while a rate limit is not.
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
    // Deliberately narrow. A bare "not found" also matches our own "IDS sheet
    // not found" wording, and a missing tab is a defect to look at, not an
    // expected outcome — only Drive's and Sheets' file-level phrasings belong
    // here.
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

  /** Whatever the thrown thing was, as text. */
  text: function (error) {
    if (!error) return "";
    if (typeof error === "string") return error;
    // `typeof`, not truthiness. An error-shaped object carrying an *empty*
    // message used to fall past this line into String(error) and come back
    // as "[object Object]" — which is then what Error Reporting grouped the
    // event on, collecting every unrelated empty-message failure into one
    // meaningless group.
    if (typeof error.message === "string") return error.message;
    if (error.message) return String(error.message);
    try {
      return String(error);
    } catch (ignored) {
      return "Unstringifiable error";
    }
  },

  /**
   * A stack trace string. Error Reporting groups on this, so when the thrown
   * value carries no stack we synthesise a single frame naming the source —
   * without it the event is dropped instead of grouped.
   *
   * Never synthesises an empty first line: the text is what the group is
   * named after, so a frame with nothing in front of it is worse than
   * useless. `summary` (the code, or whatever the caller knows) stands in
   * when there is no message at all.
   */
  stack: function (source, error, summary) {
    if (error && error.stack) return String(error.stack);
    var name = error && error.name ? error.name : "Error";
    var text = errors.text(error) || summary || source;
    return `${name}: ${text}\n    at ${source} (${source}:0:0)`;
  },

  /** Hashed so the logs can count affected users without holding addresses. */
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

  version: function () {
    try {
      return (
        PropertiesService.getScriptProperties().getProperty("APP_VERSION") ||
        "unversioned"
      );
    } catch (ignored) {
      return "unversioned";
    }
  },

  /**
   * The one shape of a bug's log entry — the payload Error Reporting ingests
   * directly. Both places that ever write one build it here (`errors.flush`
   * server-side, `reportServerError` when the browser hands a bug back), so
   * the two cannot drift into logging the same incident two different ways.
   *
   * @param {Object} fields source, code, reference, detail, trace, service,
   *                        and optionally stack, note, data.
   * @private
   */
  _event: function (fields) {
    var detail = fields.detail || "";
    var entry = {
      "@type": ERROR_REPORT_TYPE,
      // The stack's first line is what Error Reporting names the group, so
      // it must never be empty — `errors.stack` falls back to the summary.
      message:
        fields.stack ||
        errors.stack(
          fields.source,
          { message: detail },
          `${fields.code} in ${fields.source}`,
        ),
      serviceContext: {
        service: fields.service,
        version: errors.version(),
      },
      context: {
        reportLocation: { functionName: fields.source },
        user: errors.userKey(),
      },
      reference: fields.reference,
      source: fields.source,
      code: fields.code,
      expected: false,
      detail: detail,
      trace: fields.trace,
    };
    if (fields.note) entry.note = fields.note;
    if (fields.data) entry.data = fields.data;
    return entry;
  },

  /**
   * Write a bug's entry here and now, for a failure that will never reach the
   * browser to be written there.
   *
   * `record` defers a bug's write to the round trip because it cannot tell
   * whether it is the request's last stop. `reportFinal`'s call sites can:
   * they are the ones that catch a failure, log it, and then carry on as if
   * nothing happened. Nothing is returned to relay, so there is no envelope,
   * no round trip, and — before this existed — no entry at all.
   *
   * Never throws.
   *
   * @param {Object} report What `errors.report` returned.
   */
  flush: function (report) {
    try {
      // An expected outcome was already written by `record`, and only a bug
      // is ever deferred, so this is the only kind left to write.
      if (!report || report.expected) return report;
      var fingerprint = `${report.source}|${report.detail}`;
      if (_throttleReference(fingerprint, report.reference)) return report;

      errors._write(
        errors._event({
          service: "the-tower-app-script",
          source: report.source,
          code: report.code,
          reference: report.reference,
          detail: report.detail,
          trace: report.trace,
          stack: report.stack,
          note: report.note,
          data: report.data,
        }),
      );
    } catch (ignored) {
      /* reporting must not be the thing that breaks a request */
    }
    return report;
  },

  /**
   * `report` for a failure the code handles and moves on from — a cache that
   * would not open, an optional cleanup step, a dialog that never got as far
   * as the browser. Use it wherever you would have written
   * `errors.report(...)` and then ignored what it returned.
   *
   * It writes the entry immediately rather than waiting for a round trip that
   * is never coming, and then forgets the incident so a later `propagate` in
   * the same execution cannot pick it up and report some unrelated failure
   * under this one's reference.
   *
   * Same arguments as `report`.
   */
  reportFinal: function (source, error, context, code) {
    var report = errors.report(source, error, context, code);
    errors.flush(report);
    errors._forget();
    return report;
  },

  /**
   * Log one error. Returns the report so the caller can hand it to
   * `errors.fail` — never throws, whatever happens inside it.
   *
   * @param {string} source   Function the error came out of.
   * @param {*} error         The caught value.
   * @param {Object} [context] Anything that helps diagnose it — inputs,
   *                           intermediate results, IDs. Pass raw locals
   *                           directly; every value goes through
   *                           `errors.snapshot` before it is logged, so a
   *                           sheet-range array or an entire payload object is
   *                           safe to include as-is. Do not pass a raw email;
   *                           `errors.userKey()` already identifies the user.
   *                           `{ note: "..." }` is the one reserved key —
   *                           what the code was doing, in words.
   * @param {string} [code]   Override the classification.
   */
  report: function (source, error, context, code) {
    var resolvedCode = code || errors.classify(error);
    var detail = errors.text(error);
    return errors.record(source, resolvedCode, detail, error, context);
  },

  /**
   * Builds a report and, for an expected outcome, writes it immediately —
   * there is no reference to look anything up by later and no relay chain
   * worth waiting on. A bug is never written here at all. Apps Script gives
   * a function no way to tell whether it is the one `google.script.run`
   * actually invoked or just another hop in between, so instead of guessing
   * (a hand-maintained list of "boundary" function names, kept in sync by
   * hand, silently stale the moment someone forgets to update it), the write
   * waits for the one place that question answers itself: the browser has
   * it, or it doesn't. See `reportServerError` below.
   *
   * `report` and `reject` both land here so the EXPECTED/bug split is
   * applied identically to a caught exception and a precondition the code
   * checked itself. Never throws: reporting must not be the thing that
   * breaks a request.
   *
   * @param {string} source
   * @param {string} code
   * @param {string} detail   Technical text. Empty for a self-diagnosed reject.
   * @param {*} [error]       The caught value, when there was one.
   * @param {Object} [context]
   */
  record: function (source, code, detail, error, context) {
    var expected = errors.isExpected(code);
    var reference = expected ? "" : errors.reference();

    // `note` is what the code was doing, `data` is anything else the call site
    // passed. Both are omitted when absent rather than logged as empty, and
    // neither is called "details" — one character away from "detail" is not a
    // difference you can see in a log entry.
    var note = "";
    var data = null;
    if (context) {
      Object.keys(context).forEach(function (key) {
        if (key === "note") {
          if (context.note) note = context.note;
          return;
        }
        data = data || {};
        data[key] = errors.snapshot(context[key]);
      });
    }

    // A real stack, when V8 gave us one, is worth carrying all the way to
    // the eventual write — the caught value itself cannot survive the trip
    // (not serialisable, and gone once this execution ends either way), but
    // its stack string can.
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
      // Deliberately not a ReportedErrorEvent: Error Reporting is for
      // defects, and this is the app behaving correctly.
      var fields = { source: source, code: code, expected: expected };
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

  /** The one place a console call actually happens. Never throws. */
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
   * The failure envelope the client reads. Backwards compatible with the old
   * `{ success: false, message }` shape — `message` is now written for the
   * user, with the technical text kept in `detail` for the collapsible panel.
   *
   * @param {Object} report    What `errors.report` returned.
   * @param {string} [message] Overrides the message for this call site.
   * @param {Object} [extra]   Extra envelope fields (`collection: true`, …).
   */
  fail: function (report, message, extra) {
    var envelope = {
      success: false,
      code: (report && report.code) || errors.CODES.INTERNAL,
      // true when this is the app working as designed, so the client can show
      // it as a warning and leave the reference out
      expected: Boolean(report && report.expected),
      message: message || (report && report.message) || errors.MESSAGES.INTERNAL,
      reference: (report && report.reference) || "",
      detail: (report && report.detail) || "",
      // Where it actually went wrong, deepest first. Travels to the client so
      // the panel can say more than "something failed somewhere", and so
      // that if this is a bug, `reportServerError` has the complete chain
      // when the client eventually sends it back for logging.
      trace: (report && report.trace) || (report && [report.source]) || [],
    };
    // Carried along for the same reason, but only when present - an envelope
    // headed for the panel should not grow fields nobody is going to read.
    //
    // Only for a bug, though. These three exist to be handed back for the
    // deferred write, and an expected outcome has already been written
    // server-side by `record` - so sending its diagnostics out to a browser
    // that has no use for them is a round trip of internal wording and raw
    // locals for nothing.
    if (!envelope.expected) {
      if (report && report.note) envelope.note = report.note;
      if (report && report.data) envelope.data = report.data;
      if (report && report.stack) envelope.stack = report.stack;
    }

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },

  /**
   * Hand an inner failure onward without recording it a second time.
   *
   * An outer layer turning "the module said no" into its own envelope is not a
   * new incident, and reporting it as one gives Error Reporting two groups and
   * the user two references for a single root cause. Use this instead of
   * `reject` whenever the failure you are returning is somebody else's:
   *
   *     var importResult = sheetTypeFunction.importData(data, newSheetID);
   *     if (!importResult || !importResult.success) {
   *       return errors.propagate("importData", importResult, `${sheetType} could not be imported.`);
   *     }
   *
   * The reference comes from the inner envelope, or from whatever was recorded
   * deeper in this same execution — a `SheetsAPI` wrapper that reported and
   * returned null, say. If nothing was recorded at all, this really is the
   * first sighting, so it falls through to `reject`.
   *
   * @param {string} source
   * @param {Object} inner     Whatever the inner call returned.
   * @param {string} [message] Overrides the user-facing text.
   * @param {Object} [extra]   Extra envelope fields.
   */
  propagate: function (source, inner, message, extra) {
    var fresh = Date.now() - errors._lastRecordedAt < 60000;
    // Borrow the deeper failure's code as well as its reference, so an expected
    // outcome stays expected instead of being escalated to a defect on its way
    // out, and the panel titles the incident the way the log entry does.
    var code =
      (inner && inner.code) ||
      (fresh ? errors._lastCode : "") ||
      errors.CODES.INTERNAL;
    var reference =
      (inner && inner.reference) || (fresh ? errors._lastReference : "");
    // The technical text belongs to the layer that actually failed, not to
    // whatever the layer above chose to say about it.
    var detail =
      (inner && inner.detail) || (fresh ? errors._lastDetail : "") || "";
    var trace = (inner && inner.trace) ||
      (fresh && errors._lastSource ? [errors._lastSource] : []);
    var note = (inner && inner.note) || (fresh ? errors._lastNote : "") || "";
    var data = (inner && inner.data) || (fresh ? errors._lastData : null);
    var stack = (inner && inner.stack) || (fresh ? errors._lastStack : "") || "";

    if (!(inner && inner.code) && !fresh) {
      // Nothing anywhere recorded this, so we are the first to see it.
      return errors.reject(source, code, message || (inner && inner.message), extra, {
        note: inner && inner.message,
      });
    }

    // No write here, and no decision to make about whether one is due:
    // `record` never wrote this in the first place, and won't until the
    // client hands it back via `reportServerError` once it is confirmably
    // the last stop. This is a pure relay - extend the trace, pass the rest
    // through untouched.
    var envelope = {
      success: false,
      code: code,
      expected:
        inner && typeof inner.expected === "boolean"
          ? inner.expected
          : errors.isExpected(code),
      // When the code was inherited from deeper down, the caller wrote its
      // message without knowing what kind of failure this was — "could not
      // read the data" when the real answer is "Google is rate-limiting you".
      // For an expected failure the code's own message is the actionable one.
      message:
        !(inner && inner.code) && errors.isExpected(code)
          ? errors.MESSAGES[code]
          : message || (inner && inner.message) || errors.MESSAGES[code],
      reference: reference,
      detail: detail,
      trace: trace.concat(source),
    };
    // Same rule as `fail`: diagnostics ride out to the browser only because a
    // bug's entry gets written there. An expected one was written already.
    if (!envelope.expected) {
      if (note) envelope.note = note;
      if (data) envelope.data = data;
      if (stack) envelope.stack = stack;
    }

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },

  /**
   * A failure with no exception behind it — a precondition the code checked
   * itself. The code still decides how it is treated: an EXPECTED one is a
   * warning with no reference, while `SHEET_STRUCTURE` or `INVALID_INPUT`
   * means the code asked for something that should have been there, which is
   * a defect whether or not anything threw.
   *
   * @param {string} source    Function rejecting the call.
   * @param {string} code      One of errors.CODES.
   * @param {string} message   What the user is told.
   * @param {Object} [extra]   Extra envelope fields (`collection: true`, …).
   * @param {Object} [context] Diagnostics. For an **expected** outcome these
   *                           stay server-side; for a bug they ride out to
   *                           the browser and back, because that round trip
   *                           is what writes the entry. Same rules as
   *                           `report`’s context: pass raw locals, they are
   *                           snapshotted automatically.
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
      // `record`'s own `detail` is just `resolved` again here (a reject has
      // no separate exception text) - not worth repeating in the "technical
      // details" panel, so the internal note takes its place when there is
      // one, same as before this field existed as its own thing.
      //
      // The reason itself is not lost by leaving it out: it is `message`,
      // right above, and `reportServerError` falls back to `message` when
      // `detail` is empty. That fallback is load-bearing for every reject
      // without a note - which is most of them - so if the client ever stops
      // relaying `message`, they all start logging as empty entries again.
      //
      // A note is written for whoever reads the logs, in the logs' own terms
      // ("Old spreadsheet™ not found with ID: 1aBcD"). For a bug it belongs
      // in the panel, under "Technical details", because the whole point is
      // to have something to report. For an expected outcome there is nothing
      // to report and the message above already says what to do, so the note
      // stays where it was written for: the log.
      detail: report.expected ? "" : report.note || "",
      trace: report.trace,
    };
    if (!envelope.expected) {
      if (report.note) envelope.note = report.note;
      if (report.data) envelope.data = report.data;
    }

    if (extra) {
      Object.keys(extra).forEach(function (key) {
        envelope[key] = extra[key];
      });
    }

    return envelope;
  },
};

/**
 * Throttle, and — this is the point — say *which* reference the surviving
 * entry was written under.
 *
 * Shared by `errors.flush`, `reportClientError` and `reportServerError` so a
 * user hammering the same broken action, in the browser or on the server,
 * doesn't flood the log with near-identical entries.
 *
 * The reason it returns a reference rather than a boolean: the caller has
 * already minted one and, for anything coming from a page, is already showing
 * it to the user with a copy button beside it. Suppressing the write while
 * letting that reference stand hands the user an id that exists nowhere —
 * they quote it, and there is nothing to find. So the first write's reference
 * is what gets cached, and every throttled caller gets it back to show
 * instead of its own.
 *
 * @param {string} fingerprintKey
 * @param {string} reference The caller's reference, kept if this is the first.
 * @returns {string} The reference of the entry that was already written, or
 *                   "" when nothing was — meaning the caller should write.
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
    // A cache that will not open is no reason to drop the entry - better a
    // duplicate in the log than a silence.
  }
  return "";
}

/**
 * Client-side error intake. The pages call this from window.onerror, from
 * unhandled rejections, and whenever a google.script.run call fails, so
 * browser-side failures land in the same Cloud Logging stream as server ones.
 *
 * Identical errors are logged once per CLIENT_ERROR_THROTTLE_SECONDS per user
 * so a render loop cannot flood the log. A throttled call comes back with the
 * reference of the entry that *was* written, which the page shows in place of
 * its own — see `_throttleReference`.
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
        version: errors.version(),
      },
      context: {
        reportLocation: { functionName: source },
        user: errors.userKey(),
      },
      reference: reference,
      source: source,
      code: errors.CODES.CLIENT,
      detail: message,
      data: {
        page: String(safe.page || "").slice(0, 120),
        viewType: String(safe.viewType || "").slice(0, 40),
        userAgent: String(safe.userAgent || "").slice(0, 300),
        // Bounded like everything else that reaches a log entry: this comes
        // in over the wire, so its shape is whatever the page sent.
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
 * Where a bug's log entry actually gets written now. `record()` builds it
 * and hands it onward through the returned envelope — through however many
 * `propagate` hops — but never writes it, because no server function can
 * tell whether it is the one `google.script.run` actually invoked or just
 * another link in the chain. The browser can: once `AppError.show()` sees a
 * server envelope with `expected: false` and a reference, that failure has
 * definitively finished propagating (nothing server-side calls back into
 * another request to relay it further), so it calls straight back here with
 * whatever the envelope was carrying — the full accumulated `trace` among it
 * — and this is the one and only place that writes the entry.
 *
 * The trade-off against the alternative (a hand-maintained list of which
 * server functions are the real entry points, flushing there instead): this
 * requires no such list, ever, for any function added in the future — but
 * the entry now depends on this round trip actually completing. If the tab
 * closes or the connection drops between the client receiving the failure
 * and this call landing, the incident is not written at all. Weighed
 * against a stale list silently dropping entries for a function someone
 * forgot to register, this project's owner chose this trade-off
 * deliberately, aware of it.
 *
 * Same per-fingerprint throttle as `reportClientError`.
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
    // `message` is the fallback that keeps every note-less `errors.reject`
    // from logging as an empty entry - its reason lives there, not in
    // `detail`. See the comment on `reject`'s `detail` field.
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
      detail: detail,
      trace: trace,
      stack: stack,
      note: safe.note ? String(safe.note).slice(0, 500) : "",
      // This arrives over the wire. It was snapshotted on the way out, but
      // nothing guarantees what comes back is what we sent, so bound it again
      // rather than writing whatever shape the page handed us.
      data:
        safe.data && typeof safe.data === "object"
          ? errors.snapshot(safe.data)
          : null,
    });

    errors._write(entry);
    return { success: true, reference: reference };
  } catch (error) {
    console.error(`reportServerError failed: ${error && error.message}`);
    return { success: false, reference: "" };
  }
}
