#!/usr/bin/env node

/**
 * Bumps the version.
 *
 *   npm run bump patch          1.4.1 -> 1.4.2
 *   npm run bump minor          1.4.1 -> 1.5.0
 *   npm run bump major          1.4.1 -> 2.0.0
 *   npm run bump minor min      … and make this release the supported floor
 *
 * src/00_Version.js holds the version, and holds it alone: it is the only copy
 * that ships, because clasp pushes just src/. package.json deliberately has no
 * `version` field — this is not an npm package and is never published.
 *
 * The current value is read by *evaluating* that file and calling
 * appVersion.running(), not by matching text, so the answer comes from the
 * same accessor the app itself uses. Writing has to be textual: Apps Script
 * source has no filesystem, so nothing inside it can write itself. The write
 * is therefore re-evaluated afterwards to prove it produced valid JavaScript
 * carrying the intended values, and the file is put back if it did not.
 *
 * Only edits files. Nothing is committed, tagged or staged, and it does not
 * care whether the working tree is clean — the version change is left for your
 * own next commit, alongside whatever else it belongs with.
 */
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const target = path.join(__dirname, "src", "00_Version.js");
const TYPES = ["major", "minor", "patch"];

const args = process.argv.slice(2);
const type = args.find((a) => TYPES.includes(a.toLowerCase()));
const raiseMinimum = args.some((a) => /^(--?min(imum)?|min)$/i.test(a));

if (!type) {
  console.error("Usage: npm run bump <major|minor|patch> [min]");
  console.error("");
  console.error("  npm run bump patch        bump the patch version");
  console.error("  npm run bump minor min    bump the minor version and raise");
  console.error("                            the supported floor to it, so");
  console.error("                            older copies are told to update");
  console.error("");
  console.error("Edits files only — nothing is committed, tagged or staged.");
  process.exit(1);
}

/**
 * Evaluate src/00_Version.js and hand back its appVersion object.
 *
 * Needs no Apps Script stubs: the file only declares things at load time, and
 * running() reads a plain member. minimum() and latest() would need
 * PropertiesService, but they answer a different question — the floor that has
 * been *published*, not the one this release declares — so nothing here calls
 * them.
 *
 * @param {string} source
 * @returns {{VERSION: string, MINIMUM: string, running: Function}}
 * @throws {Error} When the file will not evaluate or has no appVersion.
 */
function loadAppVersion(source) {
  const sandbox = {};
  vm.createContext(sandbox);
  try {
    vm.runInContext(source, sandbox, { filename: "00_Version.js" });
  } catch (error) {
    throw new Error(`src/00_Version.js will not evaluate: ${error.message}`);
  }
  let appVersion;
  try {
    appVersion = vm.runInContext("appVersion", sandbox);
  } catch (ignored) {
    appVersion = null;
  }
  if (!appVersion || typeof appVersion.running !== "function") {
    throw new Error(
      "src/00_Version.js does not declare an appVersion object with a " +
        "running() method. It is source, not a build artefact — restore it " +
        "from git.",
    );
  }
  return appVersion;
}

/**
 * Replace one `  NAME: "…",` member, leaving the rest of the file alone.
 * @param {string} source
 * @param {string} name
 * @param {string} value
 * @returns {string}
 * @throws {Error} When the member is not there to replace.
 */
function writeMember(source, name, value) {
  const pattern = new RegExp(`^  ${name}: ".*",$`, "m");
  if (!pattern.test(source)) {
    throw new Error(`Could not find "  ${name}: …," in src/00_Version.js.`);
  }
  return source.replace(pattern, `  ${name}: ${JSON.stringify(value)},`);
}

/**
 * The next version after `current` for a major/minor/patch bump.
 * @param {string} current
 * @param {string} bump One of TYPES.
 * @returns {string}
 * @throws {Error} When current is not a plain x.y.z version.
 */
function next(current, bump) {
  const parts = /^(\d+)\.(\d+)\.(\d+)$/.exec(String(current || "").trim());
  if (!parts) {
    throw new Error(
      `appVersion.running() is "${current}", which is not a plain x.y.z ` +
        "version, so there is nothing to bump from. Fix it by hand first.",
    );
  }
  const major = Number(parts[1]);
  const minor = Number(parts[2]);
  const patch = Number(parts[3]);
  if (bump === "major") return `${major + 1}.0.0`;
  if (bump === "minor") return `${major}.${minor + 1}.0`;
  return `${major}.${minor}.${patch + 1}`;
}

try {
  if (!fs.existsSync(target)) {
    throw new Error(
      `${target} is missing. It is source, not a build artefact — restore it ` +
        "from git rather than letting this recreate it.",
    );
  }

  const before = fs.readFileSync(target, "utf8");
  const current = loadAppVersion(before);

  const from = current.running();
  const version = next(from, type.toLowerCase());
  const minimum = raiseMinimum ? version : current.MINIMUM;

  let after = writeMember(before, "VERSION", version);
  after = writeMember(after, "MINIMUM", minimum);
  fs.writeFileSync(target, after);

  // Prove the text edit produced valid JavaScript saying what was intended,
  // rather than trusting the replacement. Put the file back if it did not.
  try {
    const written = loadAppVersion(fs.readFileSync(target, "utf8"));
    if (written.running() !== version || written.MINIMUM !== minimum) {
      throw new Error(
        `wrote ${version}/${minimum} but the file reads back as ` +
          `${written.running()}/${written.MINIMUM}`,
      );
    }
  } catch (error) {
    fs.writeFileSync(target, before);
    throw new Error(
      `The bump was rolled back: ${error.message}. src/00_Version.js is ` +
        "unchanged.",
    );
  }

  console.log(
    `${from} -> ${version}` +
      (minimum ? `, minimum supported ${minimum}` : ", no minimum set") +
      (raiseMinimum ? "  (floor raised)" : ""),
  );
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
