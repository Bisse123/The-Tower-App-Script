#!/usr/bin/env node

/**
 * Checks src/ for the failures that only show up after a push.
 *
 *   npm run check
 *   npm run check -- --compare <git-ref> [--page <old>=<new> ...]
 *
 * Runs these checks over every file under src/, subfolders included:
 *
 *   1. Every .js file and every scriptlet-free <script> block parses.
 *   2. Every `object.member` used on a top-level server object resolves to a
 *      member that object declares. Quoted strings and comment lines are not
 *      read; template literals are.
 *   3. Every include()/includeTemplate() and every page name the server opens
 *      names an existing .html file.
 *   4. No two script blocks loaded by the same page declare the same global.
 *   5. Every server function the client calls, through runAppsScript or
 *      google.script.run, is a top-level server function, and so is every
 *      entry point Apps Script or CI calls by name.
 *   6. Every function an inline handler (onclick="…") calls is a global on the
 *      page that carries the handler.
 *   7. Each page's scripts run, in page order, in a stand-in browser without a
 *      ReferenceError: nothing runs before the code it needs has loaded.
 *      Scriptlets that call a server function with no arguments, such as
 *      errorContract(), print what the real server code returns.
 *   8. Every server .js file loads on its own, so no file uses another's
 *      symbols while Apps Script loads it and push order never matters.
 *
 * --compare <git-ref> also loads src/ as it was at that ref and requires every
 * page to be unchanged apart from where its code sits: the same lines, its
 * style rules in the same order, the same globals and the same outcome when
 * its scripts run. Use it after moving
 * client code, with the commit before the move as the ref. --page old=new
 * pairs a page that was renamed.
 *
 * Exits 1 when any check fails, printing each failure.
 */
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { spawnSync } = require("child_process");

const ROOT = __dirname;

const BUILTIN_MEMBERS = new Set([
  "map", "push", "forEach", "filter", "find", "findIndex", "some", "every",
  "reduce", "indexOf", "includes", "slice", "splice", "join", "concat", "sort",
  "length", "keys", "values", "entries", "hasOwnProperty", "toString",
  "apply", "call", "bind", "pop", "shift", "unshift", "at",
]);

const SERVER_ENTRY_POINTS = [
  "doGet", "onOpen", "onInstall", "include", "includeTemplate",
  "setLatestAppVersion", "getAppVersionStatus",
];

const HANDLER_NAMES = new Set([
  "event", "this", "window", "document", "console", "return", "if", "typeof",
  "true", "false", "null", "undefined",
]);

const failures = [];

/**
 * Reads every .js and .html file under src/, from the working tree or from a
 * git ref.
 * @param {string|null} ref A git ref, or null for the working tree.
 * @returns {Object<string, string>} File text by path relative to src/, with
 *   forward slashes and no byte-order mark.
 */
function loadTree(ref) {
  const files = {};
  const keep = (file) => file.endsWith(".js") || file.endsWith(".html");
  if (!ref) {
    const src = path.join(ROOT, "src");
    fs.readdirSync(src, { recursive: true })
      .map((file) => file.split(path.sep).join("/"))
      .filter(keep)
      .forEach((file) => {
        files[file] = fs.readFileSync(path.join(src, file), "utf8");
      });
  } else {
    const git = (args) => {
      const result = spawnSync("git", args, { cwd: ROOT, encoding: "utf8", maxBuffer: 1 << 28 });
      if (result.status !== 0) throw new Error(`git ${args.join(" ")}: ${result.stderr}`);
      return result.stdout;
    };
    git(["ls-tree", "-r", "--name-only", ref, "--", "src"])
      .split("\n")
      .filter((file) => file && keep(file))
      .forEach((file) => {
        files[file.replace(/^src\//, "")] = git(["show", `${ref}:${file}`]);
      });
  }
  Object.keys(files).forEach((file) => {
    files[file] = files[file].replace(/^﻿/, "").replace(/\r\n/g, "\n");
  });
  return files;
}

/**
 * The paths in a tree with one extension, sorted.
 * @param {Object<string, string>} tree
 * @param {string} extension
 * @returns {string[]}
 */
function filesWith(tree, extension) {
  return Object.keys(tree).filter((file) => file.endsWith(extension)).sort();
}

/**
 * Every <script> block in some HTML that has no src attribute.
 * @param {string} html
 * @returns {string[]} Each block's code.
 */
function scriptBlocks(html) {
  return [...html.matchAll(/<script(\s[^>]*)?>([\s\S]*?)<\/script>/g)]
    .filter((match) => !/\ssrc\s*=/.test(match[1] || ""))
    .map((match) => match[2]);
}

/**
 * Stand-ins for the Apps Script services, so server files can load outside
 * Apps Script.
 * @returns {Object} A vm context with every service name defined.
 */
function standInAppsScript() {
  const stub = new Proxy(function () {}, {
    get: (target, prop) => (prop === Symbol.toPrimitive ? () => "" : stub),
    apply: () => stub,
  });
  const sandbox = { console: { log() {}, warn() {}, error() {}, info() {} } };
  [
    "SpreadsheetApp", "DriveApp", "Drive", "Sheets", "HtmlService", "PropertiesService",
    "CacheService", "Utilities", "Session", "ScriptApp", "Logger", "UrlFetchApp",
    "LockService", "ContentService",
  ].forEach((name) => (sandbox[name] = stub));
  return vm.createContext(sandbox);
}

const serverContexts = new WeakMap();

/**
 * The server code of a tree, loaded once, for evaluating page scriptlets.
 * @param {Object<string, string>} tree
 * @returns {Object} A vm context holding every server file.
 */
function serverContext(tree) {
  if (!serverContexts.has(tree)) {
    const context = standInAppsScript();
    filesWith(tree, ".js").forEach((file) => {
      try {
        vm.runInContext(tree[file], context, { filename: file });
      } catch (ignored) {
      }
    });
    serverContexts.set(tree, context);
  }
  return serverContexts.get(tree);
}

/**
 * Replaces Apps Script scriptlets so the code can be parsed and run without
 * the server. A printing scriptlet that calls a server function with no
 * arguments, such as errorContract(), prints what that function returns;
 * any other printing scriptlet prints __STUB__, and the rest are removed.
 * @param {string} code
 * @param {Object<string, string>} tree The tree whose server code to call.
 * @returns {string}
 */
function stubScriptlets(code, tree) {
  return code
    .replace(/<\?!?=([\s\S]*?)\?>/g, (all, expression) => {
      if (!/^\s*[A-Za-z_$][\w$]*\(\)\s*$/.test(expression)) return "__STUB__";
      try {
        return String(vm.runInContext(expression, serverContext(tree)));
      } catch (error) {
        failures.push(`SCRIPTLET ${expression.trim()}: ${error.name}: ${error.message}`);
        return "__STUB__";
      }
    })
    .replace(/<\?[\s\S]*?\?>/g, "");
}

/**
 * Check 1: every .js file and every scriptlet-free <script> block parses.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkSyntax(tree) {
  filesWith(tree, ".js").forEach((file) => {
    try {
      new vm.Script(tree[file], { filename: file });
    } catch (error) {
      failures.push(`SYNTAX ${file}: ${error.message}`);
    }
  });
  filesWith(tree, ".html").forEach((file) => {
    scriptBlocks(tree[file]).forEach((code) => {
      if (code.includes("<?")) return;
      try {
        new vm.Script(code, { filename: file });
      } catch (error) {
        failures.push(`SYNTAX ${file}: ${error.message}`);
      }
    });
  });
}

/**
 * Check 2: every member used on a top-level server object is declared on it.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkMembers(tree) {
  const objects = {};
  const sources = {};
  filesWith(tree, ".js").forEach((file) => {
    const lines = tree[file].split("\n");
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
 * The top-level server functions.
 * @param {Object<string, string>} tree
 * @returns {Set<string>}
 */
function serverFunctions(tree) {
  const names = new Set();
  filesWith(tree, ".js").forEach((file) => {
    for (const match of tree[file].matchAll(/^(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)) {
      names.add(match[1]);
    }
  });
  return names;
}

/**
 * The HTML files that are whole pages rather than fragments.
 * @param {Object<string, string>} tree
 * @returns {string[]} Names without the .html extension.
 */
function pageNames(tree) {
  return filesWith(tree, ".html")
    .filter((file) => /<html[\s>]/i.test(tree[file]))
    .map((file) => file.replace(/\.html$/, ""));
}

/**
 * A page with every include inlined, as the server would serve it.
 * @param {Object<string, string>} tree
 * @param {string} name The page or fragment name, without .html.
 * @returns {Array<{source: string, text: string}>} The page's text in document
 *   order, each piece tagged with the file it came from.
 */
function expandPage(tree, name) {
  const text = tree[`${name}.html`];
  if (text === undefined) return [];
  const pieces = [];
  let last = 0;
  for (const match of text.matchAll(/<\?!=\s*include(?:Template)?\('([^']+)'\);?\s*\?>/g)) {
    pieces.push({ source: name, text: text.slice(last, match.index) });
    pieces.push(...expandPage(tree, match[1]));
    last = match.index + match[0].length;
  }
  pieces.push({ source: name, text: text.slice(last) });
  return pieces;
}

/**
 * The script blocks of an expanded page, in the order the browser runs them.
 * @param {Array<{source: string, text: string}>} pieces
 * @param {Object<string, string>} tree The tree the page came from, for its
 *   scriptlets.
 * @returns {Array<{source: string, code: string}>}
 */
function pageScripts(pieces, tree) {
  return pieces.flatMap((piece) =>
    scriptBlocks(piece.text).map((code) => ({ source: piece.source, code: stubScriptlets(code, tree) })),
  );
}

/**
 * The names a script block declares at its top level.
 * @param {string} code
 * @returns {string[]}
 */
function blockGlobals(code) {
  const lines = code.split("\n");
  const indents = lines
    .filter((line) => line.trim() !== "")
    .map((line) => line.match(/^ */)[0].length);
  if (indents.length === 0) return [];
  const top = Math.min(...indents);
  const names = [];
  lines.forEach((line) => {
    if (line.trim() === "" || line.match(/^ */)[0].length !== top) return;
    const text = line.slice(top);
    const named = text.match(/^(?:async\s+)?function\s*\*?\s*([A-Za-z_$][\w$]*)|^class\s+([A-Za-z_$][\w$]*)/);
    if (named) {
      names.push(named[1] || named[2]);
      return;
    }
    const variable = text.match(/^(?:let|const|var)\s+(.*)$/);
    if (!variable) return;
    const rest = variable[1];
    if (/^[{[]/.test(rest)) {
      const close = rest.indexOf(rest[0] === "{" ? "}" : "]");
      rest.slice(1, close).split(",").forEach((part) => {
        const name = part.split(":").pop().split("=")[0].trim();
        if (/^[A-Za-z_$][\w$]*$/.test(name)) names.push(name);
      });
      return;
    }
    const head = rest.split(/[=;]/)[0];
    head.split(",").forEach((part) => {
      const name = part.trim();
      if (/^[A-Za-z_$][\w$]*$/.test(name)) names.push(name);
    });
  });
  return names;
}

/**
 * Check 3: includes and the page names the server opens resolve.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkIncludes(tree) {
  const known = new Set(filesWith(tree, ".html").map((file) => file.replace(/\.html$/, "")));
  filesWith(tree, ".html").forEach((file) => {
    for (const match of tree[file].matchAll(/include(?:Template)?\('([^']+)'\)/g)) {
      if (!known.has(match[1])) failures.push(`INCLUDE ${file}: no file named ${match[1]}`);
    }
  });
  filesWith(tree, ".js").forEach((file) => {
    for (const match of tree[file].matchAll(/create(?:Template|HtmlOutput)FromFile\(\s*["']([^"']+)["']/g)) {
      if (!known.has(match[1])) failures.push(`PAGE ${file}: no file named ${match[1]}`);
    }
  });
}

/**
 * Check 4: no page loads the same global from two script blocks.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkDuplicateGlobals(tree) {
  pageNames(tree).forEach((page) => {
    const owners = {};
    pageScripts(expandPage(tree, page), tree).forEach((block) => {
      new Set(blockGlobals(block.code)).forEach((name) => {
        (owners[name] = owners[name] || []).push(block.source);
      });
    });
    Object.keys(owners).forEach((name) => {
      if (owners[name].length > 1) {
        failures.push(`GLOBAL ${page}: ${name} is declared in ${owners[name].join(", ")}`);
      }
    });
  });
}

/**
 * Skips a balanced parenthesised group, ignoring parentheses inside strings.
 * @param {string} text
 * @param {number} start Index of the opening parenthesis.
 * @returns {number} Index just past the matching closing parenthesis, or -1.
 */
function skipParens(text, start) {
  let depth = 0;
  let quote = null;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (quote) {
      if (ch === "\\") i++;
      else if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === "`") quote = ch;
    else if (ch === "(") depth++;
    else if (ch === ")" && --depth === 0) return i + 1;
  }
  return -1;
}

/**
 * The server functions one client file calls.
 * @param {string} text The file's text.
 * @returns {Array<{name: string|null, detail: string}>} A null name is a call
 *   whose target could not be worked out.
 */
function clientServerCalls(text) {
  const calls = [];
  for (const match of text.matchAll(/runAppsScript\(\s*(["'])([^"']+)\1/g)) {
    calls.push({ name: match[2], detail: `runAppsScript("${match[2]}")` });
  }
  for (const match of text.matchAll(/runAppsScript\(\s*(?!["'])([A-Za-z_$][\w$]*)/g)) {
    if (!/function\s+runAppsScript\s*$/.test(text.slice(0, match.index + "runAppsScript".length))) {
      calls.push({ name: null, detail: `runAppsScript(${match[1]})` });
    }
  }
  for (const match of text.matchAll(/google\.script\.run\b/g)) {
    let i = match.index + match[0].length;
    for (;;) {
      const rest = text.slice(i);
      const handler = rest.match(/^\s*\.\s*(withSuccessHandler|withFailureHandler|withUserObject)\s*\(/);
      if (!handler) break;
      i = skipParens(text, i + handler[0].length - 1);
      if (i < 0) break;
    }
    if (i < 0) continue;
    const rest = text.slice(i);
    const direct = rest.match(/^\s*\.\s*([A-Za-z_$][\w$]*)\s*\(/);
    const dynamic = rest.match(/^\s*\[\s*([A-Za-z_$][\w$]*)\s*\]\s*\(/);
    if (direct) {
      calls.push({ name: direct[1], detail: `google.script.run.${direct[1]}` });
    } else if (dynamic && dynamic[1] !== "method") {
      const declaration = text.match(
        new RegExp(`(?:const|let|var)\\s+${dynamic[1]}\\s*=([^;]+);`),
      );
      const names = declaration ? [...declaration[1].matchAll(/["']([A-Za-z_$][\w$]*)["']/g)] : [];
      if (names.length === 0) calls.push({ name: null, detail: `google.script.run[${dynamic[1]}]` });
      names.forEach((name) => calls.push({ name: name[1], detail: `google.script.run[${dynamic[1]}]` }));
    }
  }
  return calls;
}

/**
 * Check 5: every server function the client or Apps Script calls exists.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkServerCalls(tree) {
  const server = serverFunctions(tree);
  SERVER_ENTRY_POINTS.forEach((name) => {
    if (!server.has(name)) failures.push(`ENTRY ${name} is not a top-level server function`);
  });
  filesWith(tree, ".html").forEach((file) => {
    clientServerCalls(tree[file]).forEach((call) => {
      if (call.name === null) {
        failures.push(`CALL ${file}: cannot tell which server function ${call.detail} calls`);
      } else if (!server.has(call.name)) {
        failures.push(`CALL ${file}: ${call.detail} has no top-level server function`);
      }
    });
  });
}

/**
 * The globals a page declares across all its script blocks.
 * @param {Array<{source: string, code: string}>} scripts
 * @returns {Set<string>}
 */
function pageGlobals(scripts) {
  return new Set(scripts.flatMap((block) => blockGlobals(block.code)));
}

/**
 * Check 6: every name an inline handler calls is a global on its page.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkHandlers(tree) {
  pageNames(tree).forEach((page) => {
    const pieces = expandPage(tree, page);
    const globals = pageGlobals(pageScripts(pieces, tree));
    const missing = new Set();
    pieces.forEach((piece) => {
      for (const match of piece.text.matchAll(/\son[a-z]+\s*=\s*(["'])([\s\S]*?)\1/g)) {
        const handler = match[2]
          .replace(/\$\{[^}]*\}/g, "")
          .replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, "");
        for (const call of handler.matchAll(/(?<![\w$.])([A-Za-z_$][\w$]*)\s*[.(]/g)) {
          const name = call[1];
          if (HANDLER_NAMES.has(name) || globals.has(name)) continue;
          missing.add(`${piece.source}: ${name}`);
        }
      }
    });
    missing.forEach((entry) => failures.push(`HANDLER ${page}: ${entry} is not a global on this page`));
  });
}

/**
 * A stand-in browser: any property of the DOM, google or gapi is another
 * stand-in, so page code can run without a real page around it.
 * @returns {Object} A vm context.
 */
function standInBrowser() {
  const stub = new Proxy(function () {}, {
    get(target, prop) {
      if (prop === Symbol.toPrimitive) return () => "";
      if (prop === Symbol.iterator) return function* () {};
      if (prop === "then") return undefined;
      if (prop === "length") return 0;
      return stub;
    },
    apply: () => stub,
    construct: () => stub,
    set: () => true,
    has: () => false,
  });
  const noop = () => 0;
  const StandIn = function () {
    return stub;
  };
  const sandbox = {
    console: { log: noop, warn: noop, error: noop, info: noop, debug: noop },
    document: stub,
    navigator: { userAgent: "", maxTouchPoints: 0, platform: "", language: "en", clipboard: stub },
    location: stub,
    history: stub,
    screen: stub,
    google: stub,
    gapi: stub,
    localStorage: stub,
    sessionStorage: stub,
    setTimeout: noop,
    setInterval: noop,
    clearTimeout: noop,
    clearInterval: noop,
    requestAnimationFrame: noop,
    addEventListener: noop,
    removeEventListener: noop,
    dispatchEvent: noop,
    matchMedia: () => stub,
    getComputedStyle: () => stub,
    alert: noop,
    confirm: () => false,
    prompt: () => null,
    fetch: () => stub,
    open: () => stub,
    close: noop,
    focus: noop,
    innerWidth: 1024,
    innerHeight: 768,
    devicePixelRatio: 1,
    MutationObserver: StandIn,
    ResizeObserver: StandIn,
    IntersectionObserver: StandIn,
    Event: StandIn,
    CustomEvent: StandIn,
    HTMLElement: StandIn,
    Element: StandIn,
    Node: StandIn,
    FileReader: StandIn,
    Blob: StandIn,
    File: StandIn,
    Image: StandIn,
    URL,
    URLSearchParams,
    TextDecoder,
    TextEncoder,
    atob,
    btoa,
    performance: { now: () => 0 },
    __STUB__: stub,
  };
  sandbox.window = sandbox;
  sandbox.self = sandbox;
  sandbox.top = sandbox;
  sandbox.parent = sandbox;
  return vm.createContext(sandbox);
}

/**
 * Runs a page's scripts in order in a stand-in browser.
 * @param {Array<{source: string, code: string}>} scripts
 * @returns {{errors: Array<{source: string, text: string}>, defined: Set<string>}}
 *   What each failing block threw, and which of the page's globals exist once
 *   every block has run.
 */
function simulateLoad(scripts) {
  const context = standInBrowser();
  const errors = [];
  scripts.forEach((block) => {
    try {
      vm.runInContext(block.code, context, { filename: block.source, timeout: 2000 });
    } catch (error) {
      errors.push({ source: block.source, text: `${error.name}: ${error.message}` });
    }
  });
  const defined = new Set();
  pageGlobals(scripts).forEach((name) => {
    try {
      vm.runInContext(name, context);
      defined.add(name);
    } catch (ignored) {
    }
  });
  return { errors, defined };
}

/**
 * Check 7: no page hits a ReferenceError while its scripts load.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkLoad(tree) {
  pageNames(tree).forEach((page) => {
    simulateLoad(pageScripts(expandPage(tree, page), tree)).errors.forEach((error) => {
      if (error.text.startsWith("ReferenceError")) {
        failures.push(`LOAD ${page}: ${error.source} stops with ${error.text}`);
      }
    });
  });
}

/**
 * Check 8: every server file loads on its own, so none uses another file's
 * symbols while Apps Script loads it and the push order never matters.
 * @param {Object<string, string>} tree
 * @returns {void}
 */
function checkServerLoad(tree) {
  filesWith(tree, ".js").forEach((file) => {
    try {
      vm.runInContext(tree[file], standInAppsScript(), { filename: file });
    } catch (error) {
      failures.push(`SERVER LOAD ${file}: ${error.name}: ${error.message}`);
    }
  });
}

const WRAPPER_LINES = new Set(["<script>", "</script>", "<style>", "</style>"]);

/**
 * The text lines of an expanded page, leaving out include lines and the bare
 * <script>, </script>, <style> and </style> lines that splitting a fragment
 * adds.
 * @param {Array<{source: string, text: string}>} pieces
 * @returns {string[]} Trimmed, non-blank lines, sorted.
 */
function pageLines(pieces) {
  return pieces
    .flatMap((piece) => piece.text.split("\n"))
    .map((line) => line.trim())
    .filter((line) => line !== "" && !WRAPPER_LINES.has(line))
    .sort();
}

/**
 * Every style rule of an expanded page, in cascade order.
 * @param {Array<{source: string, text: string}>} pieces
 * @returns {string} The CSS of every <style> block, joined and with blank
 *   lines removed.
 */
function pageStyles(pieces) {
  return pieces
    .flatMap((piece) => [...piece.text.matchAll(/<style>([\s\S]*?)<\/style>/g)].map((m) => m[1]))
    .join("\n")
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line !== "")
    .join("\n");
}

/**
 * The differences between two sorted lists, as counted lines.
 * @param {string[]} before
 * @param {string[]} after
 * @returns {string[]} "-" for a line lost, "+" for a line gained.
 */
function listDifferences(before, after) {
  const counts = {};
  before.forEach((line) => (counts[line] = (counts[line] || 0) - 1));
  after.forEach((line) => (counts[line] = (counts[line] || 0) + 1));
  return Object.keys(counts)
    .filter((line) => counts[line] !== 0)
    .map((line) => `${counts[line] > 0 ? "+" : "-"}${Math.abs(counts[line])} ${line}`);
}

/**
 * Compare mode: every page must hold the same code as at the ref, declare the
 * same globals, and load with the same outcome.
 * @param {Object<string, string>} before The tree at the ref.
 * @param {Object<string, string>} after The working tree.
 * @param {Object<string, string>} renames Old page name to new page name.
 * @returns {void}
 */
function comparePages(before, after, renames) {
  const afterPages = new Set(pageNames(after));
  pageNames(before).forEach((oldPage) => {
    const page = renames[oldPage] || oldPage;
    if (!afterPages.has(page)) {
      failures.push(`COMPARE ${oldPage}: no page named ${page} now`);
      return;
    }
    const oldPieces = expandPage(before, oldPage);
    const newPieces = expandPage(after, page);

    listDifferences(pageLines(oldPieces), pageLines(newPieces))
      .slice(0, 20)
      .forEach((line) => failures.push(`COMPARE ${page}: line ${line}`));

    if (pageStyles(oldPieces) !== pageStyles(newPieces)) {
      failures.push(`COMPARE ${page}: its style rules no longer appear in the same order`);
    }

    const oldScripts = pageScripts(oldPieces, before);
    const newScripts = pageScripts(newPieces, after);
    listDifferences([...pageGlobals(oldScripts)].sort(), [...pageGlobals(newScripts)].sort())
      .forEach((name) => failures.push(`COMPARE ${page}: global ${name}`));

    const oldLoad = simulateLoad(oldScripts);
    const newLoad = simulateLoad(newScripts);
    listDifferences(oldLoad.errors.map((e) => e.text).sort(), newLoad.errors.map((e) => e.text).sort())
      .forEach((text) => failures.push(`COMPARE ${page}: load error ${text}`));
    listDifferences([...oldLoad.defined].sort(), [...newLoad.defined].sort())
      .forEach((name) => failures.push(`COMPARE ${page}: defined after load ${name}`));
  });
}

/**
 * Runs every check on a tree and returns what failed.
 * @param {Object<string, string>} tree
 * @param {Object<string, string>|null} baseline A tree to compare pages with,
 *   or null to skip the comparison.
 * @param {Object<string, string>} renames Old page name to new page name.
 * @returns {string[]} Each problem once.
 */
function runChecks(tree, baseline, renames) {
  failures.length = 0;
  checkSyntax(tree);
  checkMembers(tree);
  checkIncludes(tree);
  checkDuplicateGlobals(tree);
  checkServerCalls(tree);
  checkHandlers(tree);
  checkLoad(tree);
  checkServerLoad(tree);
  if (baseline) comparePages(baseline, tree, renames);
  return [...new Set(failures)];
}

if (require.main === module) {
  const args = process.argv.slice(2);
  const compareIndex = args.indexOf("--compare");
  const compareRef = compareIndex >= 0 ? args[compareIndex + 1] : null;
  const renames = {};
  args.forEach((arg, i) => {
    if (args[i - 1] === "--page") {
      const [oldPage, newPage] = arg.split("=");
      renames[oldPage] = newPage;
    }
  });

  const tree = loadTree(null);
  const problems = runChecks(tree, compareRef ? loadTree(compareRef) : null, renames);
  problems.forEach((problem) => console.log(problem));
  console.log(
    `${filesWith(tree, ".js").length} .js and ${filesWith(tree, ".html").length} .html files checked` +
      `${compareRef ? `, pages compared with ${compareRef}` : ""}, ${problems.length} problem(s).`,
  );
  process.exit(problems.length > 0 ? 1 : 0);
}

module.exports = { loadTree, runChecks, pageNames };
