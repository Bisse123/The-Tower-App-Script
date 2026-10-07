#!/usr/bin/env node

/**
 * Checks src/ for the failures that only show up after a push.
 *
 *   npm run check
 *
 * Runs four checks over every file under src/, subfolders included:
 *
 *   1. Every .js file parses (node --check).
 *   2. Every <script> block in every .html file parses. Blocks holding a
 *      scriptlet (<?) are skipped, as Apps Script evaluates those first.
 *   3. Every `object.member` used on a top-level server object resolves to a
 *      member that object declares. Quoted strings and comment lines are not
 *      read; template literals are.
 *   4. Every include()/includeTemplate() names an existing .html file, and no
 *      two script fragments loaded by the same page declare the same global.
 *
 * Exits 1 when any check fails, printing each failure.
 */
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const SRC = path.join(__dirname, "src");

const BUILTIN_MEMBERS = new Set([
  "map", "push", "forEach", "filter", "find", "findIndex", "some", "every",
  "reduce", "indexOf", "includes", "slice", "splice", "join", "concat", "sort",
  "length", "keys", "values", "entries", "hasOwnProperty", "toString",
  "apply", "call", "bind", "pop", "shift", "unshift", "at",
]);

const failures = [];

/**
 * Lists every file under a directory, as paths relative to src/ with forward
 * slashes.
 * @param {string} extension The extension to keep, e.g. ".js".
 * @returns {string[]}
 */
function listSrc(extension) {
  return fs
    .readdirSync(SRC, { recursive: true })
    .map((file) => file.split(path.sep).join("/"))
    .filter((file) => file.endsWith(extension))
    .sort();
}

/**
 * Reads a file under src/ without its byte-order mark.
 * @param {string} file Path relative to src/.
 * @returns {string}
 */
function readSrc(file) {
  return fs.readFileSync(path.join(SRC, file), "utf8").replace(/^﻿/, "");
}

/**
 * Check 1: every .js file parses.
 * @param {string[]} jsFiles
 * @returns {void}
 */
function checkJsSyntax(jsFiles) {
  jsFiles.forEach((file) => {
    const result = spawnSync(process.execPath, ["--check", path.join(SRC, file)], {
      encoding: "utf8",
    });
    if (result.status !== 0) {
      failures.push(`SYNTAX ${file}: ${(result.stderr || "").trim()}`);
    }
  });
}

/**
 * Check 2: every scriptlet-free <script> block parses.
 * @param {string[]} htmlFiles
 * @returns {void}
 */
function checkHtmlSyntax(htmlFiles) {
  htmlFiles.forEach((file) => {
    const blocks = readSrc(file).match(/<script>([\s\S]*?)<\/script>/g) || [];
    blocks.forEach((block) => {
      const code = block.replace(/^<script>/, "").replace(/<\/script>$/, "");
      if (code.includes("<?")) return;
      try {
        new Function(code);
      } catch (error) {
        failures.push(`SYNTAX ${file}: ${error.message}`);
      }
    });
  });
}

/**
 * Check 3: every member used on a top-level server object is declared on it.
 * @param {string[]} jsFiles
 * @returns {void}
 */
function checkMembers(jsFiles) {
  const objects = {};
  const sources = {};
  jsFiles.forEach((file) => {
    const lines = readSrc(file).split("\n");
    sources[file] = lines
      .filter((line) => !/^\s*(\*|\/\*|\/\/)/.test(line))
      .map((line) => line.replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, '""'))
      .join("\n");
    lines.forEach((line, index) => {
      const declaration = line.match(/^const ([A-Za-z_$][\w$]*) = \{\s*$/);
      if (!declaration) return;
      const members = new Set();
      for (let i = index + 1; i < lines.length && !/^\};?\s*$/.test(lines[i]); i++) {
        const member =
          lines[i].match(/^  (?:get |async )?([A-Za-z_$][\w$]*)\s*(?::|\()/);
        if (member) members.add(member[1]);
      }
      objects[declaration[1]] = members;
    });
  });

  const names = Object.keys(objects);
  if (names.length === 0) return;
  const usage = new RegExp(`(?<![\\w$.])(${names.join("|")})\\.([A-Za-z_$][\\w$]*)`, "g");
  Object.keys(sources).forEach((file) => {
    const seen = new Set();
    for (const match of sources[file].matchAll(usage)) {
      const [, objectName, member] = match;
      if (objects[objectName].has(member) || BUILTIN_MEMBERS.has(member)) continue;
      const key = `${objectName}.${member}`;
      if (seen.has(key)) continue;
      seen.add(key);
      failures.push(`MISSING ${file}: ${key} is not declared on ${objectName}`);
    }
  });
}

/**
 * The fragments a page includes, in include order.
 * @param {string} text The page's HTML.
 * @returns {string[]} Names as written in the include calls.
 */
function includedNames(text) {
  return [...text.matchAll(/include(?:Template)?\('([^']+)'\)/g)].map((m) => m[1]);
}

/**
 * The globals declared at the top level of every <script> block in a file.
 * @param {string} text
 * @returns {string[]}
 */
function scriptGlobals(text) {
  const globals = [];
  const blocks = text.match(/<script>([\s\S]*?)<\/script>/g) || [];
  blocks.forEach((block) => {
    const lines = block.replace(/^<script>/, "").replace(/<\/script>$/, "").split("\n");
    const indents = lines
      .filter((line) => line.trim() !== "")
      .map((line) => line.match(/^ */)[0].length);
    if (indents.length === 0) return;
    const topIndent = " ".repeat(Math.min(...indents));
    lines.forEach((line) => {
      if (!line.startsWith(topIndent) || line[topIndent.length] === " ") return;
      const declaration = line
        .slice(topIndent.length)
        .match(/^(?:async\s+)?(?:function\s+([A-Za-z_$][\w$]*)|(?:let|const|var)\s+([A-Za-z_$][\w$]*))/);
      if (declaration) globals.push(declaration[1] || declaration[2]);
    });
  });
  return globals;
}

/**
 * Check 4: includes resolve, and no page loads the same global twice.
 * @param {string[]} htmlFiles
 * @returns {void}
 */
function checkPages(htmlFiles) {
  const known = new Set(htmlFiles.map((file) => file.replace(/\.html$/, "")));
  htmlFiles.forEach((file) => {
    const text = readSrc(file);
    const included = includedNames(text);
    included.forEach((name) => {
      if (!known.has(name)) failures.push(`INCLUDE ${file}: no file named ${name}`);
    });
    if (included.length === 0) return;

    const owners = {};
    [file]
      .concat(included.filter((name) => known.has(name)).map((name) => `${name}.html`))
      .forEach((source) => {
        new Set(scriptGlobals(readSrc(source))).forEach((name) => {
          (owners[name] = owners[name] || []).push(source);
        });
      });
    Object.keys(owners).forEach((name) => {
      if (owners[name].length > 1) {
        failures.push(`GLOBAL ${file}: ${name} is declared in ${owners[name].join(", ")}`);
      }
    });
  });
}

const jsFiles = listSrc(".js");
const htmlFiles = listSrc(".html");

checkJsSyntax(jsFiles);
checkHtmlSyntax(htmlFiles);
checkMembers(jsFiles);
checkPages(htmlFiles);

failures.forEach((failure) => console.log(failure));
console.log(
  `${jsFiles.length} .js and ${htmlFiles.length} .html files checked, ${failures.length} problem(s).`,
);
process.exit(failures.length > 0 ? 1 : 0);
