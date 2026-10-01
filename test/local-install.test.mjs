import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { buildLocalSetupCommand } from "../dist/local-install.js";

const bash = process.platform === "win32" ? "C:\\Program Files\\Git\\bin\\bash.exe" : "bash";
const powershell = "powershell.exe";
const shells = ["posix", "powershell"];
const available = {
  posix: commandWorks(bash, ["--version"]),
  powershell: process.platform === "win32" && commandWorks(powershell, ["-NoProfile", "-Command", "exit 0"])
};
const defaults = { dataDir: "/selected/project", repository: "https://example.invalid/repo.git", branch: "master" };

test("defaults preserve Cowork POSIX commands and separate runtime from persistent data", () => {
  const command = buildLocalSetupCommand(defaults);
  assert.equal(command, buildLocalSetupCommand({ ...defaults, environment: "cowork", shell: "posix" }));
  assert.match(command, /^sh -eu <<'HARE_LOCAL_SETUP'/);
  assert.match(command, /HARE_RUNTIME_ROOT="\$\{XDG_CACHE_HOME:-\$HOME\/\.cache\}\/hare-m365-agent-runtime"/);
  assert.match(command, /HARE_APP="\$HARE_RUNTIME_ROOT\/app"/);
  assert.match(command, /--data-dir "\$HARE_DATA_DIR" network check --environment cowork &&\nnode/);
  assert.doesNotMatch(command, /--host|--command-shell|rm -rf|HARE_SNAPSHOT|\.hare-app-snapshot|\$HARE_DATA_DIR\/app/);
  assert.match(command, /HARE_REF="refs\/heads\/\$HARE_BRANCH"/);
  assert.doesNotMatch(command, /refs\/heads\/master/);
});

for (const shell of shells) {
  test(`${shell}: rejects invalid inputs before generating a command`, () => {
    for (const field of ["dataDir", "repository", "branch"]) {
      for (const value of ["", " ", "bad\0value", "bad\nvalue", "bad\rvalue", "bad\tvalue", "bad\x7fvalue", "--bad", null, 42]) {
        assert.throws(() => buildLocalSetupCommand({ ...defaults, shell, [field]: value }), /Hare/);
      }
    }
    for (const branch of ["-main", "../main", "feature//a", "feature/.hidden", "main.lock", "main.", "a b", "@", "@{-1}", "a~1", "a^", "a:b", "a?", "a*", "a[b", "a\\b", "a/"]) {
      assert.throws(() => buildLocalSetupCommand({ ...defaults, shell, branch }), /Hare/);
    }
    assert.throws(() => buildLocalSetupCommand({ ...defaults, shell: "cmd" }), /shell/);
    assert.throws(() => buildLocalSetupCommand({ ...defaults, environment: "codex" }), /environment/);
    if (shell === "powershell") {
      for (const field of ["dataDir", "repository", "branch"]) {
        assert.throws(() => buildLocalSetupCommand({ ...defaults, shell, [field]: 'a"b' }), /double quotes/);
      }
    }
  });

  test(`${shell}: quotes metacharacters and adds global Claude Code flags`, () => {
    const command = buildLocalSetupCommand({ ...defaults, shell, environment: "claude-code", branch: "feature/'$literal;name" });
    if (shell === "posix") {
      assert.equal(command.match(/--host claude-code --command-shell posix/g)?.length, 2);
      assert.match(command, /network check --environment claude-code/);
      assert.ok(command.includes("HARE_BRANCH='feature/'\"'\"'$literal;name'"));
    } else {
      assert.match(command, /'--host', 'claude-code', '--command-shell', 'powershell'/);
      assert.match(command, /'network', 'check', '--environment', 'claude-code'/);
      assert.ok(command.includes("$HARE_BRANCH = 'feature/''$literal;name'"));
      assert.match(command, /\$LASTEXITCODE -ne 0/);
      assert.match(command, /LOCALAPPDATA/);
      assert.match(command, /\.cache\/HareM365Agent\/runtime/);
      assert.doesNotMatch(buildLocalSetupCommand({ ...defaults, shell }), /--host|--command-shell/);
    }
  });

  for (const environment of ["cowork", "claude-code"]) {
    test(`${shell}/${environment}: install, cache reuse, update and rebuild preserve selected data`, { skip: !available[shell] }, (t) => {
      const f = fixture(t, shell, { environment, branch: "feature/'$literal;name" });
      success(f.run());
      assert.equal(f.buildCount(), 1);
      assert.equal(fs.existsSync(path.join(f.app, ".git")), true);
      assert.equal(fs.existsSync(path.join(f.dataDir, "app")), false);
      const events = f.events().filter((event) => event.kind === "network" || event.kind === "startup");
      assert.deepEqual(events.map((event) => event.kind), ["network", "startup"]);
      for (const event of events) event.args[1] = path.resolve(event.args[1]);
      const prefix = ["--data-dir", f.dataDir];
      if (environment === "claude-code") prefix.push("--host", "claude-code", "--command-shell", shell);
      assert.deepEqual(events[0].args, [...prefix, "network", "check", "--environment", environment]);
      assert.deepEqual(events[1].args, prefix);
      assert.equal(f.marker(), git(f.source, "rev-parse", "HEAD"));
      success(f.run());
      assert.equal(f.buildCount(), 1, "unchanged HEAD must not rebuild");
      f.update();
      success(f.run());
      assert.equal(f.buildCount(), 2, "new HEAD must rebuild once");
      assert.equal(f.marker(), git(f.source, "rev-parse", "HEAD"));
      fs.unlinkSync(path.join(f.app, "dist", "msal-network.js"));
      success(f.run());
      assert.equal(f.buildCount(), 3, "missing output must rebuild even at the same HEAD");
      assert.equal(f.cache(), "fixture-cache-must-survive");
    });
  }

  for (const failure of ["remote", "branch", "dirty", "ci", "build", "incomplete"]) {
    test(`${shell}: ${failure} failure never updates build marker or starts CLI`, { skip: !available[shell] }, (t) => {
      const f = fixture(t, shell);
      success(f.run());
      const oldMarker = f.marker();
      const oldHead = git(f.app, "rev-parse", "HEAD");
      f.update();
      if (failure === "remote") git(f.app, "remote", "set-url", "origin", path.join(f.root, "different-local-repo"));
      if (failure === "branch") git(f.app, "checkout", "-b", "other");
      if (failure === "dirty") fs.writeFileSync(path.join(f.app, "untracked work.txt"), "keep my changes");
      f.clearEvents();
      const result = f.run({ HARE_TEST_FAIL: failure });
      assert.notEqual(result.status, 0, result.stdout);
      assert.equal(f.marker(), oldMarker);
      assert.equal(f.cache(), "fixture-cache-must-survive");
      assert.equal(f.events().some((event) => ["network", "startup"].includes(event.kind)), false);
      if (["remote", "branch", "dirty"].includes(failure)) {
        assert.match(result.stderr, new RegExp(failure));
        assert.equal(git(f.app, "rev-parse", "HEAD"), oldHead, "must not update checkout after a failed safety check");
        assert.deepEqual(f.events(), []);
      } else {
        assert.deepEqual(f.events().map((event) => event.kind), failure === "ci" ? ["npm-ci"] : ["npm-ci", "npm-build"]);
        success(f.run());
        assert.equal(f.marker(), git(f.source, "rev-parse", "HEAD"), "retry rebuilds before recording success");
      }
      if (failure === "dirty") assert.equal(fs.readFileSync(path.join(f.app, "untracked work.txt"), "utf8"), "keep my changes");
    });
  }

  for (const failure of ["ci", "build"]) {
    test(`${shell}: first ${failure} failure leaves no successful build marker`, { skip: !available[shell] }, (t) => {
      const f = fixture(t, shell);
      assert.notEqual(f.run({ HARE_TEST_FAIL: failure }).status, 0);
      assert.equal(fs.existsSync(f.markerPath), false);
      assert.equal(f.events().some((event) => ["network", "startup"].includes(event.kind)), false);
      assert.deepEqual(f.events().map((event) => event.kind), failure === "ci" ? ["npm-ci"] : ["npm-ci", "npm-build"]);
      assert.equal(f.cache(), "fixture-cache-must-survive");
    });
  }

  test(`${shell}: preserves an existing non-Git app directory`, { skip: !available[shell] }, (t) => {
    const f = fixture(t, shell);
    fs.mkdirSync(f.app, { recursive: true });
    const existing = path.join(f.app, "keep.txt");
    fs.writeFileSync(existing, "existing user content");
    fs.writeFileSync(f.markerPath, "existing-marker");
    const result = f.run();
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /not a Git repository/);
    assert.equal(fs.readFileSync(existing, "utf8"), "existing user content");
    assert.equal(f.marker(), "existing-marker");
    assert.deepEqual(f.events(), []);
    assert.equal(f.cache(), "fixture-cache-must-survive");
  });

  test(`${shell}: missing remote branch stops before clone/build/startup`, { skip: !available[shell] }, (t) => {
    const f = fixture(t, shell);
    const result = f.run({}, { branch: "missing" });
    assert.notEqual(result.status, 0);
    assert.equal(fs.existsSync(f.app), false);
    assert.equal(fs.existsSync(f.markerPath), false);
    assert.deepEqual(f.events(), []);
  });

  test(`${shell}: network failure prevents startup, and startup failure propagates`, { skip: !available[shell] }, (t) => {
    const f = fixture(t, shell);
    const result = f.run({ HARE_TEST_FAIL: "network" });
    assert.notEqual(result.status, 0);
    assert.deepEqual(f.events().map((event) => event.kind), ["npm-ci", "npm-build", "network"]);
    f.clearEvents();
    assert.notEqual(f.run({ HARE_TEST_FAIL: "startup" }).status, 0);
    assert.deepEqual(f.events().map((event) => event.kind), ["network", "startup"]);
    assert.equal(f.cache(), "fixture-cache-must-survive");
  });

  test(`${shell}: uses home cache only when platform cache variable is absent`, { skip: !available[shell] }, (t) => {
    const f = fixture(t, shell, { fallback: true });
    success(f.run());
    assert.equal(fs.existsSync(path.join(f.app, "dist", "cli.js")), true);
    assert.equal(f.cache(), "fixture-cache-must-survive");
  });

  for (const relation of ["equal", "ancestor", "descendant", "linked-app", "linked-data"]) {
    test(`${shell}: rejects ${relation} data/runtime overlap before Git or npm`, { skip: !available[shell] }, (t) => {
      const f = fixture(t, shell);
      let selected;
      if (relation === "equal") selected = f.runtimeRoot;
      if (relation === "ancestor") selected = path.dirname(f.runtimeRoot);
      if (relation === "descendant") selected = path.join(f.app, "selected");
      if (relation === "linked-app") {
        fs.mkdirSync(f.runtimeRoot, { recursive: true });
        fs.symlinkSync(f.dataDir, f.app, process.platform === "win32" ? "junction" : "dir");
        selected = f.dataDir;
      }
      if (relation === "linked-data") {
        fs.mkdirSync(f.runtimeRoot, { recursive: true });
        selected = path.join(f.root, "selected-alias");
        fs.symlinkSync(f.runtimeRoot, selected, process.platform === "win32" ? "junction" : "dir");
      }
      const trace = path.join(f.root, "git.trace");
      const result = f.run({ GIT_TRACE: trace }, { dataDir: selected.replaceAll("\\", "/") });
      assert.notEqual(result.status, 0);
      assert.match(result.stderr, /must not overlap/);
      assert.equal(fs.existsSync(trace), false, "must reject overlap before the first Git invocation");
      assert.equal(fs.existsSync(f.markerPath), false);
      assert.deepEqual(f.events(), []);
      assert.equal(f.cache(), "fixture-cache-must-survive");
    });
  }
}

test("POSIX fail-fast wrapper remains effective when its caller uses an OR list", { skip: !available.posix }, (t) => {
  const f = fixture(t, "posix");
  const result = f.run({ HARE_TEST_FAIL: "ci" }, {}, true);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /caller-handled-failure/);
  assert.equal(fs.existsSync(f.markerPath), false);
  assert.deepEqual(f.events().map((event) => event.kind), ["npm-ci"]);
});

test("POSIX preserves literal double quotes in paths", { skip: process.platform === "win32" || !available.posix }, (t) => {
  const f = fixture(t, "posix", { suffix: ' "double quoted"' });
  success(f.run());
  assert.equal(f.cache(), "fixture-cache-must-survive");
});

test("Git Bash keeps drive-rooted backslash XDG_CACHE_HOME at its requested location", { skip: process.platform !== "win32" || !available.posix }, (t) => {
  const f = fixture(t, "posix", { runtimeBackslash: true });
  success(f.run());
  assert.equal(fs.existsSync(path.join(f.app, "dist", "cli.js")), true);
  assert.equal(f.marker(), git(f.source, "rev-parse", "HEAD"));
  assert.equal(f.cache(), "fixture-cache-must-survive");
});

function fixture(t, shell, { environment = "claude-code", branch = "master", fallback = false, runtimeBackslash = false, suffix = "" } = {}) {
  const root = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), "hare-local-install-")));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const source = path.join(root, "source 'quoted' $literal [one]" + suffix);
  const dataDir = path.join(root, "selected 'project' $literal [two]" + suffix);
  const cacheRoot = path.join(root, "runtime 'cache' $literal [three]" + suffix);
  const home = path.join(root, "home-cache" + suffix);
  const runtimeRoot = shell === "posix"
    ? path.join(fallback ? path.join(home, ".cache") : cacheRoot, "hare-m365-agent-runtime")
    : path.join(fallback ? path.join(home, ".cache") : cacheRoot, "HareM365Agent", "runtime");
  const app = path.join(runtimeRoot, "app");
  const markerPath = path.join(runtimeRoot, ".hare-app-build-head");
  const eventsPath = path.join(root, "events.jsonl");
  const bin = path.join(root, "mock-bin");
  fs.mkdirSync(source, { recursive: true });
  fs.mkdirSync(home);
  fs.mkdirSync(bin);
  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  const cacheFile = path.join(dataDir, ".cache", "fixture-cache.txt");
  fs.writeFileSync(cacheFile, "fixture-cache-must-survive");
  writeFixtureRepository(source);
  const gitEnv = {
    GIT_CONFIG_GLOBAL: path.join(root, "no-global-config"),
    GIT_CONFIG_NOSYSTEM: "1",
    GIT_CONFIG_COUNT: "0",
    GIT_TERMINAL_PROMPT: "0",
    GIT_ALLOW_PROTOCOL: "file"
  };
  const sourceGit = (...args) => git(source, ...args, { env: gitEnv });
  sourceGit("init", "-b", branch);
  sourceGit("config", "user.email", "hare-test@example.invalid");
  sourceGit("config", "user.name", "Hare Test");
  sourceGit("add", ".");
  sourceGit("commit", "-m", "initial");
  const mockNpm = path.join(bin, "npm-fixture.mjs");
  fs.writeFileSync(mockNpm, `import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
const args = process.argv.slice(2);
const kind = args[0] === 'ci' ? 'npm-ci' : 'npm-build';
fs.appendFileSync(process.env.HARE_TEST_EVENTS, JSON.stringify({ kind, args }) + '\\n');
if (process.env.HARE_TEST_FAIL === 'ci' && args[0] === 'ci') process.exit(31);
if (args[0] === 'ci') { fs.mkdirSync('node_modules', { recursive: true }); process.exit(0); }
if (args.join(' ') !== 'run build') process.exit(32);
const result = spawnSync(process.execPath, ['build.mjs'], { stdio: 'inherit' });
process.exit(result.status ?? 33);
`);
  fs.writeFileSync(path.join(bin, "npm"), '#!/bin/sh\nexec "$HARE_TEST_NODE" "$HARE_TEST_NPM" "$@"\n', { mode: 0o755 });
  fs.writeFileSync(path.join(bin, "npm.cmd"), `@echo off\r\n"${process.execPath}" "%~dp0npm-fixture.mjs" %*\r\nexit /b %errorlevel%\r\n`);
  const env = {
    ...withoutNpxNodeShim(process.env), ...gitEnv,
    HARE_TEST_EVENTS: eventsPath, HARE_TEST_NPM: shell === "posix" ? toBashPath(mockNpm) : mockNpm,
    HARE_TEST_NODE: shell === "posix" ? toBashPath(process.execPath) : process.execPath,
    HOME: home.replaceAll("\\", "/"),
    USERPROFILE: home,
    XDG_CACHE_HOME: fallback ? "" : runtimeBackslash ? cacheRoot : cacheRoot.replaceAll("\\", "/"),
    LOCALAPPDATA: fallback ? "" : cacheRoot
  };
  for (const key of Object.keys(env)) {
    if (key.toLowerCase() === "path") env[key] = `${bin}${path.delimiter}${env[key]}`;
  }
  let version = 0;
  return {
    root, source, dataDir, runtimeRoot, app, markerPath,
    marker: () => fs.readFileSync(markerPath, "utf8").trim(),
    cache: () => fs.readFileSync(cacheFile, "utf8"),
    buildCount: () => Number(fs.readFileSync(path.join(runtimeRoot, ".fixture-build-count"), "utf8")),
    events: () => fs.existsSync(eventsPath) ? fs.readFileSync(eventsPath, "utf8").trim().split("\n").filter(Boolean).map((line) => JSON.parse(line)) : [],
    clearEvents: () => fs.writeFileSync(eventsPath, ""),
    update() {
      fs.appendFileSync(path.join(source, "build.mjs"), `\n// update ${++version}\n`);
      sourceGit("add", "build.mjs");
      sourceGit("commit", "-m", `update ${version}`);
    },
    run(extraEnv = {}, options = {}, wrapOr = false) {
      const command = buildLocalSetupCommand({
        dataDir: shell === "posix" ? dataDir.replaceAll("\\", "/") : dataDir,
        repository: process.platform === "win32" ? source.replaceAll("\\", "/") : source,
        branch, environment, shell, ...options
      });
      if (shell === "powershell") {
        return spawnSync(powershell, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", Buffer.from(command, "utf16le").toString("base64")], {
          encoding: "utf8", timeout: 120_000, env: { ...env, ...extraEnv }
        });
      }
      const body = wrapOr ? `{\n${command}\n} || printf '%s\\n' caller-handled-failure` : command;
      return spawnSync(bash, ["--noprofile", "--norc", "-c", `export PATH=${quotePosix(toBashPath(bin))}:/mingw64/bin:/usr/bin:/bin:$PATH\n${body}`], {
        encoding: "utf8", timeout: 120_000, env: { ...env, ...extraEnv }
      });
    }
  };
}

function writeFixtureRepository(source) {
  fs.writeFileSync(path.join(source, "package.json"), JSON.stringify({ name: "hare-local-install-fixture", version: "1.0.0", type: "module" }));
  fs.writeFileSync(path.join(source, ".gitignore"), "dist/\nnode_modules/\n");
  const cli = `import fs from 'node:fs';
const args = process.argv.slice(2);
const kind = args.includes('network') ? 'network' : 'startup';
fs.appendFileSync(process.env.HARE_TEST_EVENTS, JSON.stringify({ kind, args }) + '\\n');
if (process.env.HARE_TEST_FAIL === kind) process.exit(35);
`;
  fs.writeFileSync(path.join(source, "build.mjs"), `import fs from 'node:fs';
import path from 'node:path';
const countFile = path.resolve('..', '.fixture-build-count');
const count = Number(fs.existsSync(countFile) ? fs.readFileSync(countFile, 'utf8') : '0') + 1;
fs.writeFileSync(countFile, String(count));
if (process.env.HARE_TEST_FAIL === 'build') process.exit(34);
fs.mkdirSync('node_modules', { recursive: true });
fs.mkdirSync('dist', { recursive: true });
fs.writeFileSync('dist/cli.js', ${JSON.stringify(cli)});
fs.writeFileSync('dist/proxy.js', 'export {};\\n');
if (process.env.HARE_TEST_FAIL === 'incomplete') {
  fs.rmSync('dist/msal-network.js', { force: true });
} else {
  fs.writeFileSync('dist/msal-network.js', 'export {};\\n');
}
`);
}

function success(result) {
  assert.equal(result.status, 0, `${result.error ?? ""}\n${result.stdout}\n${result.stderr}`);
}

function commandWorks(command, args) {
  return spawnSync(command, args, { encoding: "utf8", timeout: 10_000 }).status === 0;
}

function quotePosix(value) {
  return `'${value.replaceAll("'", `'"'"'`)}'`;
}

function toBashPath(value) {
  return process.platform === "win32" ? value.replaceAll("\\", "/").replace(/^([A-Za-z]):/, (_, drive) => `/${drive.toLowerCase()}`) : value;
}

function git(cwd, ...args) {
  const options = typeof args.at(-1) === "object" ? args.pop() : {};
  const result = spawnSync("git", args, { cwd, encoding: "utf8", env: { ...process.env, ...options.env } });
  assert.equal(result.status, 0, result.stderr);
  return result.stdout.trim();
}

function withoutNpxNodeShim(sourceEnv) {
  const env = { ...sourceEnv };
  for (const key of Object.keys(env)) {
    if (key.toLowerCase() !== "path" || !env[key]) continue;
    env[key] = env[key].split(path.delimiter).filter((entry) => !/[\\/]npm-cache[\\/]_npx[\\/]/i.test(entry)).join(path.delimiter);
  }
  return env;
}
