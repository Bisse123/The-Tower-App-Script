/**
 * Calls a top-level function in the production Apps Script project through the
 * Apps Script API, and fails loudly if it does not return.
 *
 * Usage: node run-function.js [functionName]      (default: setLatestAppVersion)
 *
 * Environment:
 *   CLIENT_ID / CLIENT_SECRET / REFRESH_TOKEN
 *     A credential minted from an OAuth client in the *same* GCP project as the
 *     script, carrying every scope src/appsscript.json declares. The clasp
 *     login used for pushing and deploying is a different client with different
 *     scopes and will be rejected here. See documentation/07-deployment.md.
 *   SCRIPT_ID
 *     Optional; falls back to the scriptId in .clasp.prod.json.
 *
 * Runs in devMode, so the call executes HEAD rather than a pinned API
 * executable version. In CI that is the code the deploy step just published,
 * because compare-src.js has already proven HEAD equals this commit.
 */
const fs = require("fs");
const path = require("path");

const ANNOTATION = process.env.GITHUB_ACTIONS ? "::error::" : "";

/**
 * The production script id, from the environment or the checked-in config.
 * @returns {string}
 * @throws {Error} When neither source provides one.
 */
function scriptId() {
  if (process.env.SCRIPT_ID) return process.env.SCRIPT_ID;
  const config = path.join(__dirname, "..", "..", ".clasp.prod.json");
  if (fs.existsSync(config)) {
    const id = JSON.parse(fs.readFileSync(config, "utf8")).scriptId;
    if (id) return id;
  }
  throw new Error(
    "No script id: set SCRIPT_ID, or restore .clasp.prod.json.",
  );
}

/**
 * Trades the refresh token for a short-lived access token.
 * @returns {Promise<string>}
 * @throws {Error} When any credential is missing or Google refuses the trade.
 */
async function accessToken() {
  const missing = ["CLIENT_ID", "CLIENT_SECRET", "REFRESH_TOKEN"].filter(
    (name) => !process.env[name],
  );
  if (missing.length) {
    throw new Error(`Missing credentials: ${missing.join(", ")}.`);
  }

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    body: new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      refresh_token: process.env.REFRESH_TOKEN,
      grant_type: "refresh_token",
    }),
  });
  const body = await response.json();

  if (!body.access_token) {
    throw new Error(
      `Could not refresh the access token: ${body.error || response.status}` +
        (body.error_description ? ` — ${body.error_description}` : "") +
        ". The refresh token may have been revoked, or belong to the wrong " +
        "OAuth client.",
    );
  }
  return body.access_token;
}

/**
 * Runs one function and hands back whatever it returned.
 * @param {string} name Top-level function name.
 * @returns {Promise<*>}
 * @throws {Error} When the API refuses the call or the function throws.
 */
async function run(name) {
  const token = await accessToken();
  const response = await fetch(
    `https://script.googleapis.com/v1/scripts/${scriptId()}:run`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ function: name, devMode: true }),
    },
  );
  const body = await response.json();

  if (body.error) {
    const detail = (body.error.details || [])[0] || {};
    if (detail.errorMessage) {
      throw new Error(`${name} threw: ${detail.errorMessage}`);
    }
    throw new Error(
      `The API refused the call: ${body.error.message} ` +
        `(${body.error.status || response.status}). Check that the script has ` +
        "an API executable deployment, that the Apps Script API is enabled, " +
        "and that the token carries every scope in src/appsscript.json.",
    );
  }

  if (!body.response) {
    throw new Error(`Unexpected API response: ${JSON.stringify(body)}`);
  }
  return body.response.result;
}

const functionName = process.argv[2] || "setLatestAppVersion";

run(functionName)
  .then((result) => {
    console.log(`${functionName} returned ${JSON.stringify(result)}`);
  })
  .catch((error) => {
    console.error(`${ANNOTATION}${error.message}`);
    process.exit(1);
  });
