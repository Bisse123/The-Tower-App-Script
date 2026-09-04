# 08 — Error handling & reporting

One path from a thrown exception to a line in Cloud Logging, a group in Error
Reporting, and a message the user can act on. The first thing it decides is
whether the failure is a defect at all — everything else follows from that.

- [Two kinds of failure](#two-kinds-of-failure)
- [The reference id](#the-reference-id)
- [Backend](#backend)
- [The envelope](#the-envelope)
- [Frontend](#frontend)
- [Google Cloud](#google-cloud)
- [Runbook](#runbook)
- [Conventions](#conventions)

---

## Two kinds of failure

The first question about any failure is whether it is **ours**. Everything else
follows from the answer, and the answer is the error code — nothing else.

| | EXPECTED | BUG |
| --- | --- | --- |
| What it is | The app working as designed on input it cannot accept | Something we got wrong |
| Examples | Sheet too old to convert · file never granted · Google rate-limiting the account · a mistyped link | A range that should have been there · an unclassified exception · a browser crash |
| Severity | `WARNING` | `ERROR` |
| Error Reporting | **No.** It is not a defect | Yes, with a stack |
| Reference shown | **No** | Yes, with a copy button |
| Panel | ⚠️ amber | ⛔ red |

Asking someone to report a working system wastes their time and ours, so an
expected failure gets no reference at all — there is nothing to look up. The
codes on the expected side live in one object, [`errors.EXPECTED`](../src/00_Errors.js):

```javascript
EXPECTED: {
  ACCESS_DENIED, INVALID_LINK, NOT_FOUND, QUOTA, TIMEOUT, VERSION_OUTDATED
}
```

Both entry points — `errors.report` for a caught exception and `errors.reject`
for a precondition — funnel into `errors.record`, which applies that table. A
quota exception caught in a try/catch is therefore treated exactly like a quota
condition the code detected itself: neither is a bug.

### Telling a bad range from a quota

Sheets throws the same `GoogleJsonResponseException` for both, but the text
differs and that is enough:

| Failure | What comes back | Code | Kind |
| --- | --- | --- | --- |
| Range/tab does not exist | 400 `Unable to parse range: Master Sheet fail` | `SHEET_STRUCTURE` | **bug** |
| Read quota | 429 `Quota exceeded for quota metric …` | `QUOTA` | expected |
| Per-user rate limit | 429 `User rate limit exceeded` | `QUOTA` | expected |
| Apps Script daily cap | `Service invoked too many times for one day` | `QUOTA` | expected |

So asking for a range that is not there is reported with a stack and a
reference; hitting a limit is not.

> The classifier is deliberately narrow about `NOT_FOUND`. A bare "not found"
> also matches our own "IDS sheet not found" wording, and a missing tab is a
> defect to look at — only Drive's and Sheets' file-level phrasings count.

### One incident, one entry, with the full trace

A failure travels up through layers, and each layer used to turn it into one of
its own: `SheetsAPI.batchGetValues` reported the real exception, then
`importData` reported *"Could not read required data from spreadsheet"* on top
of it. Two ERROR entries, two Error Reporting groups, two references — one bug.

The first fix was `errors.propagate(source, inner, message?, extra?)`, which
hands the inner failure onward without recording a second incident — the
original `report`/`reject` call already wrote the entry that matters, so
relaying it is a pure, silent transformation of the same in-memory value:

```javascript
var importResult = sheetTypeFunction.importData(data, newSheetID);
if (!importResult || !importResult.success) {
  return errors.propagate("importData", importResult, `${sheetType} could not be imported.`);
}
```

That got Cloud Logging down to one entry — but the one entry was written
*before* propagation happened, so its own `trace` field only ever named the
deepest frame. The full chain reached the browser panel (through the returned
envelope, which `propagate` keeps extending) but not Logs Explorer, because a
Cloud Logging entry cannot be edited after it is written — there is no API
for it, and Apps Script's `console.error`/`console.warn` do not expose enough
entry identity (no custom `insertId`) to fake one either. The only way to get
the complete trace *into* the entry is to not write it until the trace is
known to be complete.

The obvious next move was a registry of "boundary" functions — the ~36 the
client actually calls via `google.script.run` — so a bug written at one of
those, rather than at the deep function where it was first caught, would carry
the whole chain. That version shipped, worked, and was reverted: Apps Script
gives a function no way to ask "was I the one the client actually invoked, or
just another hop in between," so the registry could only ever be a hand-kept
guess at the answer — accurate the day it was written, silently stale the
first time someone added a new client-callable function and forgot to list it
there too.

**What ships instead**: `record()` never writes a bug's entry at all — not at
the deep catch, not anywhere server-side. The failure rides the ordinary
return value out to the browser exactly as it already did (gaining a frame at
every `propagate` hop on the way, unchanged), and the *browser* — which is
unambiguously the end of the line, since nothing server-side calls back into
another request to relay a failure further — is where `AppError.show()`
notices a bug it has not already seen and calls `reportServerError` with
everything the envelope was carrying. That is the one and only place the
entry is written, and by then the trace is not merely complete — it can even
name the client-side function that finally received it, one frame past
anything a server-side registry could ever have known about:

```
jsonPayload.trace = ["SheetsAPI.batchGetValues", "importData", "importAllData"]
                                                                 ^ the client's own caller
```

`reject`'s fresh (non-propagated) rejections go through the same `record()`,
so they get the same treatment automatically: nothing is written until the
envelope reaches the browser, same as a caught exception.

**The trade-off, chosen deliberately over the registry's**: writing the entry
now depends on that final round trip actually landing. Today, an entry is
written server-side before the response is even sent — a dropped connection
on the way back to the browser still leaves the log entry intact, the user
just never sees the panel that would have told them to look for it. Under
this design, if the tab closes or the connection drops between the browser
receiving the failure and `reportServerError` landing, the incident is not
written at all — not "written with a shorter trace," not written. Weighed
against a stale registry silently dropping every failure from a function
someone forgot to register — a mistake with no error of its own to notice it
by — this project's owner chose the round trip, with this written down as the
reason.

### The failures no round trip can carry

Deferring the write assumes the failure is *going* somewhere: back through the
return value, out to the browser, into `reportServerError`. Some are not. A
`catch` that logs and then carries on — a cache that would not open, an
optional cleanup step, a dialog that failed before it could be shown — returns
no envelope at all, so there is nothing to hand back and nobody to hand it to.
Under the deferred write those produced **no entry anywhere**.

`errors.reportFinal(source, error, context?, code?)` is `report` for exactly
that case: same arguments, but it writes the entry on the spot, because for
these there is no round trip coming and never was. Use it wherever you would
have written `errors.report(...)` and then ignored what it returned.

It also **clears `errors._last*`** afterwards, and that half matters just as
much. `propagate` reads those when the value it is relaying carries no code of
its own — but a failure the code deliberately swallowed is finished, and
leaving it there meant the next unrelated `propagate` in the same execution
inherited its reference, trace, note and data. A Laboratory version mismatch
would be filed, complete with reference, as a Cards parser crash.

The one thing that must **not** become `reportFinal` is a `SheetsAPI`/`shared`
wrapper that reports and returns `null` for its caller to relay. That `null`
*is* the handoff, and `propagate` needs the `_last*` state `reportFinal` would
wipe.

`propagate` borrows the deeper failure’s **code**, **detail** and **trace** as well as
its reference. The code is what keeps a quota a quota: the deep layer records a
WARNING and no reference, and without it the outer layer would see "no
reference" and escalate a rate limit into a defect on its way out.

When the code is inherited that way, the outer layer’s message was written
without knowing what kind of failure it was — *"Cards: could not read the
data"* when the real answer is *"Google is rate-limiting you"*. So for an
**expected** failure the code’s own message wins; for a bug the call site’s
message is kept, because it names the sheet and the technical detail is right
underneath it.

If nothing was recorded anywhere, the failure really is new, and `propagate`
falls through to `reject`. **Use `reject` only for a failure you are the first
to notice; use `propagate` for one you are passing on.**

Nothing about *adding* a new client-callable function requires touching
`00_Errors.js` at all now — that was the entire point.

Every terminal failure list — the update workflow’s per-sheet rows, the
save-file import summary, the Get Started copy results — raises the panel once
for the first failure that is **not** expected, and stays silent when every
failure in the batch was expected. The list is the record of what happened to
each sheet; the panel is the one thing worth reporting.

---

## The reference id

Every **bug** gets one: `TWR-M4X2K9-A7F3`. It is generated where the failure is
first recorded, written into the log entry, and shown to the user in the error
panel with a copy button. A screenshot is enough to find the entry:

```
jsonPayload.reference="TWR-M4X2K9-A7F3"
```

That is the whole point of the system. Everything below exists to make sure
that id survives from the throw site to the user's screen.

---

## Backend

[00_Errors.js](../src/00_Errors.js) holds it all. All 264 catch
blocks in the project route through it — the ones in the update/get-started
workflows, plus [02_SavedFile.js](../src/02_SavedFile.js)'s save-file parser:
`parseSaveFileBytes` guards its own ungzip/NRBF decoding, and — unlike a
regular multi-step workflow — does **not** abort on the first category that
fails: each of the 11 `parse*Data` functions is independent, so one category
breaking on an unexpected shape must not cost the user the other ten. Every
category's result is checked; successes go into `parsed`, and each failure is
relayed (not re-reported — see [propagate](#one-incident-one-entry-with-the-full-trace)) into
a `failedCategories` list carrying its own code/reference/detail/trace. The
client renders the successful categories exactly as before and lists the
failed ones separately, raising the error panel once for whichever failure
in that list is an actual defect. A malformed save file therefore degrades
one category at a time instead of failing the whole import, while still
getting the same reference-and-stack treatment as everything else instead of
an unreported crash:

```javascript
} catch (error) {
  var errorReport = errors.report("lab.exportData", error, { sheetType: sheetType });
  return errors.fail(errorReport);
}
```

`errors.report(source, error, context)`

- classifies the exception into a code (see below) by matching the Google API
  wording, so a mechanically converted catch block still yields a message worth
  reading;
- for an **expected** outcome, writes it immediately: `console.warn`, no
  stack, no `@type`, so Error Reporting never sees it — there is no relay
  chain worth waiting on and no reference to key a later write on anyway;
- for a **bug**, writes nothing yet — see
  [One incident, one entry, with the full trace](#one-incident-one-entry-with-the-full-trace)
  for where that actually happens;
- returns `{ reference, code, expected, message, detail, source, trace, note?, data? }`,
  where `reference` is empty for anything expected.

`errors.reportFinal(source, error, context?, code?)` is the same call for a
failure that stops here — one the code logs and then carries on from, so no
envelope is returned and no browser will ever hand it back. It writes the
entry immediately and closes the incident. See
[The failures no round trip can carry](#the-failures-no-round-trip-can-carry).

`errors.fail(report, message?, extra?)` turns that into the client envelope.
Pass `message` where the call site knows something better than the classifier
does; pass `extra` for the envelope fields some callers add (`collection: true`,
`accessible: false`, …).

`errors.reject(source, code, message, extra?, context?)` is for a failure with
**no exception behind it** — a precondition the code checked itself. It is not
automatically a warning: the code decides, exactly as it does for `report`. A
`VERSION_OUTDATED` rejection is expected; a `SHEET_STRUCTURE` one means the
code looked for something that should have been there, which is a defect
whether or not anything threw, so it gets a stack and a reference.

### Codes

| Code | Kind | Raised when | What the user is told |
| --- | :-: | --- | --- |
| `ACCESS_DENIED` | expected | Drive/Sheets refused, or a file was never granted | Grant access and try again |
| `NOT_FOUND` | expected | The file is gone, or was never shared | The sheet could not be opened |
| `INVALID_LINK` | expected | What the user typed is not a sheet link or ID | Check the link and try again |
| `INVALID_FILE` | expected | The picked file is not a valid `playerInfo.dat` — not even valid gzip | Check you picked the right file |
| `QUOTA` | expected | Google is rate-limiting the account | Wait and retry |
| `TIMEOUT` | expected | Execution time exceeded | Try fewer sheets at once |
| `VERSION_OUTDATED` | expected | The sheet is too old for this template | What the call site says — see below |
| `INVALID_INPUT` | **bug** | A required parameter never arrived — the client should not have sent that | Reload and try again |
| `SHEET_STRUCTURE` | **bug** | A tab or label the code scans for is not there | The script could not find something it needs |
| `CLIENT` | **bug** | Reported from the browser | Something went wrong in the page |
| `INTERNAL` | **bug** | Anything unclassified | Something went wrong on our side |

`INVALID_INPUT` and `SHEET_STRUCTURE` sit on the bug side deliberately. Both
can be caused by a user editing their sheet, but both are also exactly how a
regression in our own label scanning shows up, and only the log can tell those
apart — so we look.

The panel title for `SHEET_STRUCTURE` is *"Could not read part of your sheet"*,
not *"your sheet is not laid out as expected"*. The code fires both when a user
edited their sheet and when our own scanner asks for a range that does not
exist, and the panel cannot tell those apart. A title that assigns blame will
therefore be wrong some of the time, and wrong in the direction that makes the
user hunt through a sheet that was fine.

### Capturing what led to the failure

`context` is not limited to a `note`. Pass the function's own parameters, and
whatever it had computed by the time it broke — for a `var`-declared local,
that is well-defined even when the throw happened before the assignment ever
ran, because `var` hoists to the top of the function: an unset local reads back
as `undefined`, which is itself the diagnosis ("never got that far"), not a
crash.

```javascript
getVersion1_0LabLevels: function (oldLabLevelsValues) {
  try {
    var oldLabLevels = {};
    var oldLabMax = {};
    oldLabLevelsValues.forEach(function (row) { /* … */ });
    return { success: true, oldLabLevels: oldLabLevels, oldLabMax: oldLabMax };
  } catch (error) {
    var errorReport = errors.report("lab.getVersion1_0LabLevels", error, {
      oldLabLevelsValues: oldLabLevelsValues, // the input
      oldLabLevels: oldLabLevels,             // built so far
      oldLabMax: oldLabMax,                   // undefined if it never got here
    });
    return errors.fail(errorReport);
  }
},
```

Every value passed this way goes through `errors.snapshot()` before it is
logged:

- arrays and objects are capped in both breadth (10 array items, 25 object
  keys) and depth (3 levels) — a 2000-row sheet range comes back as its first
  10 rows and a count of how many were cut, not a multi-megabyte log entry;
- strings are capped at 300 characters;
- a value that points back at itself becomes `"[Circular]"` instead of
  recursing forever;
- nothing in here can throw — a value that resists serialising becomes
  `"[unserializable]"` rather than losing the whole log entry over it.

That is what makes it safe to pass raw locals without a second thought. All
232 catch blocks across the sheet modules were given their enclosing
function's own parameters this way — [03_Laboratory.js](../src/03_Laboratory.js)
additionally has hand-picked intermediate state (the worked example above),
since deciding which local mid-computation is worth keeping needs a look at
the function; the mechanical pass could only reach for parameters.

### Never tell the user to update their sheet

Updating sheets is what this app **is**. "Try updating the sheet" in the update
workflow tells someone to do the thing they are already doing, and it is the
easiest sentence in the world to write by accident when a version check fails.

So `MESSAGES.VERSION_OUTDATED` says only what is true — *"That sheet is not a
version this step can work with"* — and every call site that knows more says it
itself:

```javascript
`Your ${sheetType} is version ${oldVersion}, which this script cannot convert to the new template.`
`Your ${sheetType} is version ${oldVersion}, which this script cannot read.`
```

The save-file workflow is the one place where "update it first" is genuine
advice, because that workflow is not the updater — and it already gives that
advice without raising anything, through `renderSaveFileOutdatedSheets` and
`renderSaveFileCollectionOutdated`. An out-of-date sheet there is a blocked
import with an explanation, not an error.

The client switches on the **code**, never on message text, so rewording a
message can never change behaviour.

---

## The envelope

Every server function still returns what it always did, with three fields added:

```javascript
{
  success: false,
  code: "SHEET_STRUCTURE",        // one of the codes above
  expected: false,                // true = the app working as designed
  message: "Cards: Could not read required data from spreadsheet",
  reference: "TWR-M4X2K9-A7F3",   // matches the log entry; "" when expected
  detail: "API call to sheets.spreadsheets.values.batchGet failed with error: Unable to parse range: Master Sheet fail",
  trace: ["SheetsAPI.batchGetValues", "importData"],
}
```

`detail` and `trace` belong to the layer that **actually failed**, not to
whichever layer wrote the message. `propagate` carries both outward untouched,
so the panel can show where a failure started rather than where it surfaced:

```
▼ Technical details
   where: SheetsAPI.batchGetValues  →  importData  →  importAllData
   code:  SHEET_STRUCTURE

   API call to sheets.spreadsheets.values.batchGet failed with error:
   Unable to parse range: Master Sheet fail
```

The client appends its own frame to `trace` on the way in, so the path runs
from the failing call all the way out to the page. An expected failure shows
none of this — no trace, no code, no reference; there is no defect to locate.
What it can still show is `detail`, when that is the raw text Google itself
gave us (*"Quota exceeded for quota metric…"*), which is worth reading. What
it never shows is an internal `note`: that is written for the log, in the
log's terms, and an expected `reject` keeps it there.

`message` is written for the person in the sidebar. `detail` is the technical
text, shown only behind *Technical details* in the panel. **No server function
throws across the boundary** — that part of the old convention is unchanged.

---

## Frontend

[22_error_scripts.html](../src/22_error_scripts.html) is the one place a page
puts an error, and [22_error_section.html](../src/22_error_section.html) is
where it lands: a persistent panel with the message, a collapsible technical
detail, and the copyable reference. The single-line `#statusText` gets a short
summary, but it is no longer where errors live — the next status update used to
wipe them out.

```javascript
AppError.show(result, { source: "importData" });                // a server envelope
AppError.show(error, { source: "copyTemplates", message: "…" }); // an exception
AppError.log(error, "checkVersion");   // record it, do not interrupt the user
AppError.check(result, "importData");  // show if failed; returns true when it did
```

- `AppError.normalize` takes an envelope, a `google.script.run` transport
  failure, an `Error`, or a string, and produces one shape.
- `runAppsScript(method, ...args)` — now defined once, here — rejects with that
  normalised shape, so no page has to guess.
- `window.onerror` and `unhandledrejection` are wired in, so a browser
  exception is no longer invisible.
- `setStatusWithSpinner` clears the panel: starting new work retires the last
  failure.
- Expected failures render amber, with no reference row and no invitation to
  report anything; bugs render red with the reference and a copy button. The
  panel takes the server’s `expected` flag when there is one and falls back to
  its own copy of the code list.

The per-file lists in the update workflow are the one place a failure appears
without the panel. Those rows never show a reference either — they are a record
of what happened to each sheet, and a defect gets the panel on top.

`AppError.surfaceBatch(entries, { source })` is what those lists call. It puts
up the panel for the **first** failure that is a bug, and files an entry for
**every** one of them. The second half is not optional: a bug's entry is only
written when the browser hands it back, so a failure the batch skips over is a
failure nobody ever hears about — "not worth its own panel" and "not worth
knowing about" are different questions. It takes `{ envelope, sheetType, error }`,
the shape the workflows already build for their own lists, and the `envelope` is
the load-bearing field: without it there is no code, no reference and nothing
to report.

Anything raised in the browser is sent to `reportClientError` on the server, so
it lands in the **same Cloud Logging stream** as backend errors under
`serviceContext.service = "the-tower-app-script-client"`. Identical errors are
sent once per page load, and again at most once per 5 minutes per user, so a
render loop cannot flood the log.

**A throttled call writes nothing, and says so.** It returns the reference of
the entry that *was* written, and the page swaps that into the panel in place
of the one it minted. Otherwise the second person to hit the same bug within
five minutes is shown an id, invited to quote it, and there is nothing behind
it — the throttle would be quietly manufacturing dead references.

A server failure with **no code at all** — one of the plain
`{ success: false, message }` returns still left in the backend — classifies as
`INTERNAL`, which puts it on the bug side. The page mints a reference for it
and files it like any other bug. That is the right call (a failure we cannot
even name is exactly the kind worth looking at) but the entry can only say what
the message said. Giving those returns a real `errors.reject`/`errors.fail` is
what turns them into something diagnosable.

A **server** bug takes a different, related path: `AppError.show()` treats a
server envelope with `expected: false` and a reference as one it has not yet
seen written, and calls `AppError.finalize()`, which sends it to
`reportServerError` — a *different* server function from `reportClientError`,
because this is not a browser exception; it is the server's own bug, and the
entry belongs under `serviceContext.service = "the-tower-app-script"`, not the
client one. This is the write that `record()` deferred — see
[One incident, one entry, with the full trace](#one-incident-one-entry-with-the-full-trace)
for the whole reasoning. Same per-page-load and per-5-minute throttling as
`reportClientError`, and the same fire-and-forget shape: a failure to report
must never surface anything of its own.

---

## Google Cloud

Both script projects are attached to standard GCP projects, which is what makes
any of this visible:

| | Dev / sandbox | Production |
| --- | --- | --- |
| GCP project | `832137601831` | `1031925368251` |

### Why the entries reach Error Reporting

Whichever of `record` (an expected outcome), `errors.flush` (a bug that stops
server-side, via `reportFinal`) or `reportServerError` (a bug,
once the round trip lands) actually performs the write, it is the same shape
either way — the payload Error Reporting ingests directly:

```javascript
{
  "@type": "type.googleapis.com/google.devtools.clouderrorreporting.v1beta1.ReportedErrorEvent",
  message: "<the stack trace>",
  serviceContext: { service: "the-tower-app-script", version: "<APP_VERSION>" },
  context: { reportLocation: { functionName: "SheetsAPI.batchGetValues" }, user: "<hashed>" },
  reference: "TWR-…",
  source: "SheetsAPI.batchGetValues",  // trace[0] - the deepest frame, not necessarily who wrote it
  code: "SHEET_STRUCTURE",
  expected: false,
  detail: "API call to sheets.spreadsheets.values.batchGet failed with error: Unable to parse range: Master Sheet fail",
  trace: ["SheetsAPI.batchGetValues", "importData", "importAllData"],
  note: "Error getting spreadsheet",   // only when the call site passed one
  data: { sheetID: "1aBcD" },          // only when the call site passed one
}
```

| Field | What it is |
| --- | --- |
| `detail` | The technical text — the exception message, or for a `reject` the reason itself. Always present. |
| `note` | What the code was *doing*, in words, when the call site knows. Omitted otherwise. |
| `data` | Anything else the call site passed: IDs, sheet types. Omitted otherwise. |

`note` and `data` are omitted rather than logged empty, so an entry never
carries a field that says nothing. They are deliberately not one field called
`details`: one character away from `detail` is not a difference you can see when
you are reading a log entry at speed.

Three things are load-bearing:

1. **Severity ERROR.** `console.error` in Apps Script maps to it; `console.log`
   does not, which is why nothing reached Error Reporting before.
2. **A stack trace in `message`.** Events without one are dropped.
3. **`serviceContext.version`**, read from the `APP_VERSION` script property, so
   a regression can be attributed to a release. Set it when you cut one; it
   falls back to `"unversioned"`.

### One-time setup per project

1. Enable the **Error Reporting API** (Sheets v4 and Drive v3 are already on).
2. Grant whoever is on call *Error Reporting Viewer* and *Logs Viewer*.
3. Create a log-based metric — Logging ▸ Log-based metrics ▸ Create:
   - Name `app_script_errors`, type Counter
   - Filter: `severity>=ERROR AND jsonPayload.code!=""`
   - Label `code` from `jsonPayload.code`, label `source` from `jsonPayload.source`
4. Alert on it — Monitoring ▸ Alerting ▸ Create policy: `app_script_errors`
   grouped by `code`, above ~10 in 5 minutes. `QUOTA` and `ACCESS_DENIED` spikes
   are user-driven; `INTERNAL` spikes are ours.

### Privacy

Raw locals are fine to pass in `context` — see
[Capturing what led to the failure](#capturing-what-led-to-the-failure) —
because `errors.snapshot()` runs over every value before it is logged and caps
it in breadth, depth and length, so a sheet's worth of data cannot balloon a
log entry or leak wholesale. That covers size; it does not decide what belongs
in the log in the first place:

- Never pass a raw email. The user is already identified by an MD5 of their
  address, truncated to 12 hex characters (`errors.userKey()`) — enough to
  count affected users, not enough to identify one.
- Don't reach for something you know is sensitive — a password, an OAuth
  token, a payment detail — just because it happened to be a variable in
  scope. Nothing in this app currently holds any of those, but the rule
  outlives the current code.
- Cloud Logging access is scoped by the GCP project's IAM, same as everything
  else in it — anyone with *Logs Viewer* on the project can read what lands
  here, so the access list on the project is the actual privacy boundary, not
  what any one call site chooses to log.
- For a bug — and **only** for a bug — `data`/`note`/`detail` also transit
  through the browser on their way to being logged, because that round trip is
  what writes the entry; see
  [One incident, one entry, with the full trace](#one-incident-one-entry-with-the-full-trace).
  Nothing new is exposed to the person already looking at the page (it is
  their own request that failed), but the same "would I mind this sitting in
  a browser's memory and a network request, however briefly" question now
  applies to `context` that used to stay server-side end to end.
- An **expected** outcome's `note` and `data` never leave the server. `record`
  has already written that entry, so there is nothing for a round trip to do,
  and a note is written in the logs' own terms — *"Old spreadsheet™ not found
  with ID: 1aBcD"* — which is not what the panel should be showing someone who
  simply needs to grant access.

---

## Runbook

A bug's `timestamp` is when the browser's round trip landed, not when the
failure actually happened server-side — usually a fraction of a second later,
rarely enough to matter. What can matter: a bug whose round trip never
completes (closed tab, dropped connection) has no entry to find at all, by
design — see
[One incident, one entry, with the full trace](#one-incident-one-entry-with-the-full-trace).
Treat "no entries in `severity>=ERROR AND jsonPayload.expected=false` for a
given window" as "no bugs *the round trip could confirm*", not as a hard
guarantee. The exception is anything logged through `errors.reportFinal`,
which is written server-side as it happens and does not depend on the tab
still being open.

Two more things that make an entry not appear, both deliberate:

- **The throttle.** Identical failures (same source, same detail) are written
  once per 5 minutes per user. The later ones are not lost — they are *the
  same entry*, and the panel shows that entry's reference, so a user quoting
  it lands on the one write that happened.
- **A page that never came back.** Unchanged from above: the round trip is
  what writes a bug.

**A user reports a problem and quotes a reference**

```
jsonPayload.reference="TWR-M4X2K9-A7F3"
```

One entry, with the stack, the source function, the code, and the context the
call site passed.

**How often is this failing, and to how many people**

```
severity>=ERROR AND jsonPayload.source="collection.importData"
```

Then group by `jsonPayload.context.user` for the number of distinct users.

**What broke in the last release**

```
severity>=ERROR AND jsonPayload.serviceContext.version="4.2.4"
```

**Browser-side failures only**

```
jsonPayload.serviceContext.service="the-tower-app-script-client"
```

**Expected failures, which never reach Error Reporting** — how often people hit
a sheet too old to convert, or run into a quota:

```
severity=WARNING AND jsonPayload.expected=true
```

Worth watching as product signal rather than as a bug queue: a spike in
`VERSION_OUTDATED` means a lot of people are stuck on an old template.

**Bugs only**, which is what Error Reporting shows:

```
severity>=ERROR AND jsonPayload.expected=false
```

---

## Conventions

When you write new code:

| Situation | Use |
| --- | --- |
| `catch` around anything | `errors.report(source, error, context)` then `errors.fail(report)` |
| What to put in `context` | The function's own parameters, always. Add a mid-computation local when you can see it would narrow down where things went wrong. Pass the raw value — `errors.snapshot()` bounds it. |
| A precondition you checked yourself | `errors.reject(source, code, message)` |
| An inner call already failed | `errors.propagate(source, inner, message?)` — never `reject`, or one incident is recorded twice |
| Deciding the code | Ask whether *we* would have to change something. If yes it is a bug; if the user has to act, it is expected. Add it to `errors.EXPECTED` and to the mirror in `22_error_scripts.html`. |
| Something recovered on its own | `errors.reportFinal(...)` and carry on — do not return a failure. Not `report`: nothing is going to hand this one back to be written. |
| A wrapper that reports and returns `null` for its caller to relay | `errors.report(...)` — the `null` is the handoff, and `reportFinal` would clear the state `propagate` reads |
| A failure with no code, from old `{ success: false, message }` code | Give it one. Until then the client treats it as `INTERNAL`, mints a reference and logs it, which is right but tells you nothing about what broke |
| Client: a failed envelope | `AppError.show(result, { source })` |
| Client: a caught exception | `AppError.show(error, { source, message })` |
| Client: a failure the user need not see | `AppError.log(error, source)` |
| Client: a list of per-sheet failures | `AppError.surfaceBatch(entries, { source })` — panel for the first bug, an entry for every one of them |

`source` is `functionName` for a top-level function and `module.method` for a
sheet-module method (`lab.exportData`, `collection.importData`) — 15 modules
share those method names, and the qualifier is the only thing that tells the log
which one failed.

Do not write `console.log` for an error. It is INFO severity, it has no stack,
and nothing will ever alert on it.
