/**
 * Compares two directories of Apps Script sources for functional equality.
 *
 * Usage: node compare-src.js <dirA> <dirB>
 *   dirA = the committed source on the branch being deployed (e.g. "src")
 *   dirB = the project's current HEAD, pulled via clasp (e.g. "remote_head")
 *
 * Exits 0 when the two are identical, 1 when they differ (printing what
 * differs). Only files clasp actually pushes are compared (.js/.gs/.html/.json).
 *
 * Normalisation avoids false mismatches from cosmetic differences that clasp
 * introduces on pull:
 *   - CRLF vs LF line endings
 *   - trailing whitespace / trailing newlines
 *   - JSON key ordering / indentation (compared structurally)
 */
const fs = require("fs");
const path = require("path");

const MANAGED_EXTS = new Set([".js", ".gs", ".html", ".json"]);

function listManagedFiles(dir) {
  const files = {};
  function walk(current, rel) {
    if (!fs.existsSync(current)) return;
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const abs = path.join(current, entry.name);
      const relPath = rel ? `${rel}/${entry.name}` : entry.name;
      if (entry.isDirectory()) {
        walk(abs, relPath);
      } else if (MANAGED_EXTS.has(path.extname(entry.name).toLowerCase())) {
        files[relPath] = fs.readFileSync(abs, "utf8");
      }
    }
  }
  walk(dir, "");
  return files;
}

function sortKeys(value) {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value && typeof value === "object") {
    return Object.keys(value)
      .sort()
      .reduce((acc, key) => {
        acc[key] = sortKeys(value[key]);
        return acc;
      }, {});
  }
  return value;
}

function normalize(relPath, content) {
  if (relPath.toLowerCase().endsWith(".json")) {
    try {
      return JSON.stringify(sortKeys(JSON.parse(content)));
    } catch (_) {
      // fall through to text normalisation if it isn't valid JSON
    }
  }
  return content
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+$/gm, "")
    .replace(/\s+$/, "");
}

const [dirA, dirB] = process.argv.slice(2);
if (!dirA || !dirB) {
  console.error("Usage: node compare-src.js <dirA> <dirB>");
  process.exit(2);
}

const a = listManagedFiles(dirA);
const b = listManagedFiles(dirB);
const allNames = new Set([...Object.keys(a), ...Object.keys(b)]);

const differences = [];
for (const name of [...allNames].sort()) {
  const inA = Object.prototype.hasOwnProperty.call(a, name);
  const inB = Object.prototype.hasOwnProperty.call(b, name);
  if (!inA) {
    differences.push(`  + only on project HEAD: ${name}`);
  } else if (!inB) {
    differences.push(`  - only on branch:       ${name}`);
  } else if (normalize(name, a[name]) !== normalize(name, b[name])) {
    differences.push(`  ~ differs:              ${name}`);
  }
}

if (differences.length > 0) {
  console.error("Branch source and project HEAD are NOT identical:");
  console.error(differences.join("\n"));
  process.exit(1);
}

console.log("Branch source matches the project HEAD exactly.");
process.exit(0);
